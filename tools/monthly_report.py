#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""月度業務分析報告產生器（業務／接單導向）。

用法:
    python3 tools/monthly_report.py --month 2026-08
    python3 tools/monthly_report.py --month 2026-08 --narrative narrative.txt

讀取 data/latest.xlsx 的「資料總表」，產出:
    reports/YYYY-MM.html   自足式報告（可分享）
    reports/YYYY-MM.json   全部計算結果
    reports/index.html     歷月報告索引

資料規則與 index.html 的 buildDashboardData 一致，門檻集中在 tools/report_config.json。
分析取向為業務量（鍋數為主、半成品重量 kg 交叉驗證）；製成率／品質不在本報告範圍。
"""
import argparse
import json
import os
import re
import statistics
import sys
import zipfile
from datetime import date, datetime
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(ROOT, 'data', 'latest.xlsx')
SHEET = '資料總表'
REPORTS = os.path.join(ROOT, 'reports')
CONFIG = os.path.join(ROOT, 'tools', 'report_config.json')

M_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# 料號為這些值時代表沒有正式料號，改用「料號__品項」當 key（與儀表板一致）
PLACEHOLDER_CODES = {'-', '專案', '研發', ''}

# 欄位索引（0-based）
C_DATE, C_CODE, C_NAME, C_POTS, C_SPEC, C_KG, C_DELETED = 0, 1, 2, 3, 4, 8, 20


# ────────────────────────── xlsx 讀取（純標準函式庫） ──────────────────────────
def _col_index(ref):
    m = re.match(r'([A-Z]+)', ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_sheet(path, sheet_name):
    """回傳 list[dict[col_index] = value]（value 為 str 或 float）。"""
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels}
    target = None
    for s in wb.find('{%s}sheets' % M_NS):
        if s.get('name') == sheet_name:
            target = relmap[s.get('{%s}id' % R_NS)].lstrip('/')
    if target is None:
        raise SystemExit('找不到工作表「%s」，請確認 %s 是正確的檔案。' % (sheet_name, path))
    path_in_zip = target if target.startswith('xl/') else 'xl/' + target

    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root:
            shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % M_NS)))

    data = z.read(path_in_zip).decode('utf-8', 'ignore')
    rows = []
    for rw in re.findall(r'<row[^>]*>(.*?)</row>', data, re.S):
        cells = {}
        for attrs, inner in re.findall(r'<c ([^>]*?)(?:/>|>(.*?)</c>)', rw, re.S):
            ref = re.search(r'r="([A-Z]+\d+)"', attrs)
            if not ref:
                continue
            idx = _col_index(ref.group(1))
            t = re.search(r't="([^"]+)"', attrs)
            ttype = t.group(1) if t else None
            val = None
            v = re.search(r'<v>(.*?)</v>', inner or '', re.S)
            if v is not None:
                raw = v.group(1)
                if ttype == 's':
                    val = shared[int(raw)]
                elif ttype in (None, 'n'):
                    try:
                        val = float(raw)
                    except ValueError:
                        val = raw
                else:
                    val = raw
            else:
                istr = re.search(r'<is>.*?<t[^>]*>(.*?)</t>', inner or '', re.S)
                if istr:
                    val = istr.group(1)
            if val is not None and val != '':
                cells[idx] = val
        rows.append(cells)
    return rows


def parse_ymd(v):
    """支援字串日期與 Excel 序列值。"""
    if v is None:
        return None
    if isinstance(v, float) or isinstance(v, int):
        # Excel 序列值（1899-12-30 為基準）
        try:
            d = date.fromordinal(date(1899, 12, 30).toordinal() + int(v))
            return (d.year, d.month, d.day)
        except Exception:
            return None
    s = str(v).strip().split(' ')[0]
    for sep in ('/', '-'):
        if sep in s:
            parts = s.split(sep)
            if len(parts) != 3:
                return None
            try:
                return (int(parts[0]), int(parts[1]), int(parts[2]))
            except ValueError:
                return None
    return None


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').strip()
    if s == '':
        return None
    try:
        return float(s)
    except ValueError:
        return None


def month_add(y, m, delta):
    t = y * 12 + (m - 1) + delta
    return t // 12, t % 12 + 1


def mlabel(y, m):
    return '%d/%02d' % (y, m)


def pct(cur, base):
    """變化百分比；base 為 0 或 None 時回傳 None。"""
    if not base:
        return None
    return (cur - base) / base * 100.0


# ────────────────────────── 彙總 ──────────────────────────
def aggregate(rows):
    pots, kg, meta = {}, {}, {}
    month_counts = {}
    for r in rows[1:]:
        ymd = parse_ymd(r.get(C_DATE))
        if not ymd:
            continue
        y, m, _ = ymd
        name = r.get(C_NAME)
        if name is None:
            continue
        ns = str(name).strip()
        if ns in ('#N/A', ''):
            continue
        ym = (y, m)
        month_counts[ym] = month_counts.get(ym, 0) + 1

        code = r.get(C_CODE)
        cs = '' if code is None else str(code).strip()
        if isinstance(code, float) and code == int(code):
            cs = str(int(code))
        key = cs if cs not in PLACEHOLDER_CODES else cs + '__' + str(name)
        if key not in meta:
            meta[key] = {
                '料號': cs,
                '品項': str(name),
                '類別': '花生類' if '花生' in str(name) else '果醬餡類',
            }
        if r.get(C_DELETED):
            continue
        pv = to_num(r.get(C_POTS))
        kv = to_num(r.get(C_KG))
        if pv is not None:
            pots.setdefault(key, {})[ym] = pots.setdefault(key, {}).get(ym, 0.0) + pv
        if kv is not None:
            kg.setdefault(key, {})[ym] = kg.setdefault(key, {}).get(ym, 0.0) + kv
    return pots, kg, meta, month_counts


# ────────────────────────── 分析 ──────────────────────────
def compute(pots, kg, meta, month_counts, Y, M, cfg):
    P = lambda k, ym: pots.get(k, {}).get(ym, 0.0)
    K = lambda k, ym: kg.get(k, {}).get(ym, 0.0)

    cur = (Y, M)
    prev = month_add(Y, M, -1)
    yoy = (Y - 1, M)
    keys = sorted(meta.keys())

    # ── 完整性檢查 ──
    # 沿用儀表板規則，但只有「最新月」才可能是尚未補齊（其餘月份已收單完畢）。
    # 歷史月份的筆數本來就會隨工作天數／淡旺季浮動 ±20%，不應誤判為不完整，
    # 因此僅在明顯異常（< 中位數 50%）時才提醒。
    all_ym = sorted(month_counts.keys())
    cur_count = month_counts.get(cur, 0)
    latest_ym = all_ym[-1]
    prior = [month_counts[ym] for ym in all_ym if ym != latest_ym]
    prior_median = statistics.median(prior) if prior else 0
    latest_incomplete = bool(prior_median and month_counts[latest_ym] < 0.9 * prior_median)
    ref = [month_counts[ym] for ym in all_ym
           if ym != cur and not (ym == latest_ym and latest_incomplete)]
    median = statistics.median(ref) if ref else 0
    is_latest = (cur == latest_ym)
    if is_latest and latest_incomplete:
        incomplete = True
    else:
        incomplete = bool(median and cur_count < 0.5 * median)

    # ── 整體 ──
    tot_pots = sum(P(k, cur) for k in keys)
    tot_pots_prev = sum(P(k, prev) for k in keys)
    tot_pots_yoy = sum(P(k, yoy) for k in keys)
    tot_kg = sum(K(k, cur) for k in keys)
    tot_kg_prev = sum(K(k, prev) for k in keys)
    tot_kg_yoy = sum(K(k, yoy) for k in keys)
    active_items = [k for k in keys if P(k, cur) > 0]

    # YTD（1月~M月）今年 vs 去年同期
    ytd = sum(P(k, (Y, m)) for k in keys for m in range(1, M + 1))
    ytd_prev = sum(P(k, (Y - 1, m)) for k in keys for m in range(1, M + 1))
    ytd_kg = sum(K(k, (Y, m)) for k in keys for m in range(1, M + 1))
    ytd_kg_prev = sum(K(k, (Y - 1, m)) for k in keys for m in range(1, M + 1))

    # ── 主力品項：以「前 12 個月累計鍋數」排名（不受本月衰退影響）──
    base_window = [month_add(Y, M, -i) for i in range(1, 13)]
    baseline = {k: sum(P(k, ym) for ym in base_window) for k in keys}
    # ── 品項註記（tools/report_config.json 的 item_overrides，以料號為 key）──
    # discontinued=已下市：完全排除警示與燈號；seasonal_oem=季節性代工：不列主力、不觸發燈號
    overrides = cfg.get('item_overrides', {}) or {}
    ov = {}           # item key -> override dict
    for k in keys:
        o = overrides.get(meta[k]['料號'])
        if o:
            ov[k] = o
    dropped_keys = {k for k, o in ov.items() if o.get('status') == 'discontinued'}
    oem_keys = {k for k, o in ov.items() if o.get('status') == 'seasonal_oem'}

    # 主力品項排名時排除季節性代工與已下市品項
    rank_pool = [k for k in keys if k not in oem_keys and k not in dropped_keys]
    core_keys = set(sorted(rank_pool, key=lambda k: -baseline[k])[:cfg['top_n']])

    dec = float(cfg['decline_pct'])
    scale = float(cfg['min_scale_pots'])
    kg_flat = float(cfg['kg_flat_pct'])

    # ── 接單下滑警示 ──
    alerts = []
    for k in keys:
        if k in dropped_keys:
            continue
        pc, pp, py = P(k, cur), P(k, prev), P(k, yoy)
        kc, kp, ky = K(k, cur), K(k, prev), K(k, yoy)
        mom_pct = pct(pc, pp)
        yoy_pct = pct(pc, py)
        trig_mom = pp >= scale and mom_pct is not None and mom_pct <= -dec
        trig_yoy = py >= scale and yoy_pct is not None and yoy_pct <= -dec
        if not (trig_mom or trig_yoy):
            continue
        loss_mom = (pp - pc) if trig_mom else 0.0
        loss_yoy = (py - pc) if trig_yoy else 0.0
        # 以流失較大的基準做交叉判讀
        if loss_yoy > loss_mom:
            basis, loss, p_pct = 'YoY', loss_yoy, yoy_pct
            kg_base, kg_pct_v = ky, pct(kc, ky)
        else:
            basis, loss, p_pct = 'MoM', loss_mom, mom_pct
            kg_base, kg_pct_v = kp, pct(kc, kp)

        if kg_base <= 0 or kg_pct_v is None:
            verdict, vlabel = 'unknown', '無重量資料，無法交叉驗證'
        elif kg_pct_v <= -dec:
            verdict, vlabel = 'real', '實質接單下滑（產出同步下降）'
        elif kg_pct_v <= -kg_flat:
            verdict, vlabel = 'partial', '部分下滑'
        else:
            verdict, vlabel = 'batch', '批量調整，非接單下滑'

        alerts.append({
            'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'], '類別': meta[k]['類別'],
            'pots_cur': round(pc), 'pots_prev': round(pp), 'pots_yoy': round(py),
            'kg_cur': round(kc), 'kg_prev': round(kp), 'kg_yoy': round(ky),
            'mom_pct': None if mom_pct is None else round(mom_pct, 1),
            'yoy_pct': None if yoy_pct is None else round(yoy_pct, 1),
            'kg_pct': None if kg_pct_v is None else round(kg_pct_v, 1),
            'basis': basis, 'loss': round(loss), 'pots_pct': None if p_pct is None else round(p_pct, 1),
            'verdict': verdict, 'verdict_label': vlabel,
            'core': k in core_keys,
            'tag': ov[k]['label'] if k in ov else None,
            'oem': k in oem_keys,
        })
    # ── 斷單（前 N 個月常態生產、本月掛零）──
    # 斷單比「衰退」更嚴重，獨立成一類；同樣套用規模門檻，濾除零星小量品項的雜訊。
    look = int(cfg['dormant_lookback'])
    need = int(cfg['dormant_min_active'])
    dormant = []
    for k in keys:
        if k in dropped_keys or P(k, cur) > 0:
            continue
        window = [month_add(Y, M, -i) for i in range(1, look + 1)]
        vals = [P(k, ym) for ym in window]
        active = sum(1 for v in vals if v > 0)
        if active < need:
            continue
        avg = sum(vals) / max(1, active)
        if avg < scale:
            continue
        dormant.append({
            'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'],
            'active_months': active, 'avg_pots': round(avg),
            'last_pots': round(vals[0]), 'core': k in core_keys,
            'tag': ov[k]['label'] if k in ov else None,
        })
    dormant.sort(key=lambda d: -d['avg_pots'])
    dormant_keys = {d['key'] for d in dormant}

    # 已列為斷單者不再重複列入衰退警示，讓兩份清單互斥、計數不重複
    alerts = [a for a in alerts if a['key'] not in dormant_keys]
    alerts.sort(key=lambda a: -a['loss'])
    real_alerts = [a for a in alerts if a['verdict'] in ('real', 'partial')]
    core_real = [a for a in real_alerts if a['core']]

    # ── 成長品項 ──
    growth = []
    for k in keys:
        pc, pp, py = P(k, cur), P(k, prev), P(k, yoy)
        g_mom, g_yoy = pc - pp, pc - py
        gain = max(g_mom, g_yoy)
        if gain <= 0 or pc <= 0:
            continue
        growth.append({
            'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'],
            'pots_cur': round(pc), 'pots_prev': round(pp), 'pots_yoy': round(py),
            'gain_mom': round(g_mom), 'gain_yoy': round(g_yoy), 'gain': round(gain),
            'mom_pct': None if pct(pc, pp) is None else round(pct(pc, pp), 1),
            'yoy_pct': None if pct(pc, py) is None else round(pct(pc, py), 1),
            'tag': ov[k]['label'] if k in ov else None,
        })
    growth.sort(key=lambda g: -g['gain'])

    # ── 新品 / 回流 ──
    newbies, returning = [], []
    for k in active_items:
        hist = [ym for ym in pots.get(k, {}) if ym < cur and pots[k][ym] > 0]
        if not hist:
            newbies.append({'料號': meta[k]['料號'], '品項': meta[k]['品項'], 'pots': round(P(k, cur))})
            continue
        last3 = [month_add(Y, M, -i) for i in (1, 2, 3)]
        if all(P(k, ym) == 0 for ym in last3):
            gap_end = max(hist)
            returning.append({'料號': meta[k]['料號'], '品項': meta[k]['品項'],
                              'pots': round(P(k, cur)), 'last_seen': mlabel(*gap_end)})

    # ── 集中度 ──
    ranked = sorted(active_items, key=lambda k: -P(k, cur))[:cfg['top_n']]
    top_list, cum = [], 0.0
    for k in ranked:
        v = P(k, cur)
        cum += v
        top_list.append({'品項': meta[k]['品項'], '料號': meta[k]['料號'],
                         'pots': round(v),
                         'share': round(v / tot_pots * 100, 1) if tot_pots else 0,
                         'cum_share': round(cum / tot_pots * 100, 1) if tot_pots else 0})
    top_share = round(cum / tot_pots * 100, 1) if tot_pots else 0

    # ── 類別 ──
    cats = {}
    for cat in ('果醬餡類', '花生類'):
        ks = [k for k in keys if meta[k]['類別'] == cat]
        c, p, yv = (sum(P(k, cur) for k in ks), sum(P(k, prev) for k in ks), sum(P(k, yoy) for k in ks))
        cats[cat] = {
            'pots': round(c), 'prev': round(p), 'yoy': round(yv),
            'kg': round(sum(K(k, cur) for k in ks)),
            'mom_pct': None if pct(c, p) is None else round(pct(c, p), 1),
            'yoy_pct': None if pct(c, yv) is None else round(pct(c, yv), 1),
            'share': round(c / tot_pots * 100, 1) if tot_pots else 0,
        }

    # ── 季節性：同月跨年度，以及「該月轉換的常態變化」──
    years = sorted({y for (y, _m) in month_counts})
    same_month = []
    for y in years:
        tp = sum(P(k, (y, M)) for k in keys)
        if tp > 0 or (y, M) in month_counts:
            same_month.append({'year': y, 'pots': round(tp)})
    typical_moms = []
    for y in years:
        if y >= Y:
            continue
        py_, pm_ = month_add(y, M, -1)
        a = sum(P(k, (y, M)) for k in keys)
        b = sum(P(k, (py_, pm_)) for k in keys)
        r = pct(a, b)
        if r is not None and b > 0 and a > 0:
            typical_moms.append(r)
    typical_mom = round(sum(typical_moms) / len(typical_moms), 1) if typical_moms else None
    mom_pct_total = pct(tot_pots, tot_pots_prev)
    seasonal_note = None
    if typical_mom is not None and mom_pct_total is not None:
        gap = mom_pct_total - typical_mom
        if mom_pct_total < 0 and gap >= -5:
            seasonal_note = '本月環比 %.1f%%，過去同月份平均為 %.1f%%，變化落在季節性常態範圍內。' % (mom_pct_total, typical_mom)
        elif gap < -5:
            seasonal_note = '本月環比 %.1f%%，明顯低於過去同月份平均的 %.1f%%，非單純季節性因素。' % (mom_pct_total, typical_mom)
        else:
            seasonal_note = '本月環比 %.1f%%，優於過去同月份平均的 %.1f%%。' % (mom_pct_total, typical_mom)

    # ── 近 N 個月趨勢 ──
    trend = []
    for i in range(int(cfg['trend_months']) - 1, -1, -1):
        ym = month_add(Y, M, -i)
        if ym not in month_counts:
            continue
        trend.append({'label': mlabel(*ym),
                      'pots': round(sum(P(k, ym) for k in keys)),
                      'kg': round(sum(K(k, ym) for k in keys))})

    # ── 排除季節性代工後的總量（反映本業真實表現）──
    # 代工品項量體大且間歇，會讓月度總數劇烈起伏；另計一組排除後的數字。
    has_oem = bool(oem_keys) and any(
        P(k, cur) > 0 or P(k, prev) > 0 or P(k, yoy) > 0 for k in oem_keys)
    ex_pots = ex_prev = ex_yoy_v = ex_mom = ex_yoy_pct = None
    if has_oem:
        ex_pots = sum(P(k, cur) for k in keys if k not in oem_keys)
        ex_prev = sum(P(k, prev) for k in keys if k not in oem_keys)
        ex_yoy_v = sum(P(k, yoy) for k in keys if k not in oem_keys)
        ex_mom = pct(ex_pots, ex_prev)
        ex_yoy_pct = pct(ex_pots, ex_yoy_v)

    # ── 綜合燈號 ──
    yoy_pct_total = pct(tot_pots, tot_pots_yoy)
    reasons = []
    # 有季節性代工品項時，改用排除後的同比判定，避免代工的有無誤觸紅燈
    judge_yoy = ex_yoy_pct if has_oem else yoy_pct_total
    if judge_yoy is not None and judge_yoy <= -dec:
        reasons.append('總鍋數較去年同月下降 %.1f%%%s' % (
            judge_yoy, '（已排除季節性代工）' if has_oem else ''))
    if core_real:
        reasons.append('%d 項主力品項出現實質衰退' % len(core_real))
    core_dormant = [d for d in dormant if d['core']]
    if core_dormant:
        reasons.append('%d 項主力品項本月斷單' % len(core_dormant))
    if reasons:
        light, light_label = 'red', '需注意'
    elif real_alerts or dormant:
        light, light_label = 'yellow', '觀察'
        reasons.append('有 %d 項品項出現衰退訊號，惟未涉及主力品項' % (len(real_alerts) + len(dormant)))
    elif (mom_pct_total or 0) > 0 and (yoy_pct_total or 0) > 0:
        light, light_label = 'green', '表現良好'
        reasons.append('總量環比與同比皆成長，且無主力品項衰退警示')
    else:
        light, light_label = 'yellow', '持平'
        reasons.append('總量變化不大，無明顯衰退警示')

    return {
        'month': mlabel(Y, M), 'year': Y, 'mon': M,
        'prev_label': mlabel(*prev), 'yoy_label': mlabel(*yoy),
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'incomplete': incomplete, 'is_latest_month': is_latest,
        'record_count': cur_count, 'median_count': median,
        'totals': {
            'pots': round(tot_pots), 'pots_prev': round(tot_pots_prev), 'pots_yoy': round(tot_pots_yoy),
            'kg': round(tot_kg), 'kg_prev': round(tot_kg_prev), 'kg_yoy': round(tot_kg_yoy),
            'items': len(active_items),
            'mom_pct': None if mom_pct_total is None else round(mom_pct_total, 1),
            'yoy_pct': None if yoy_pct_total is None else round(yoy_pct_total, 1),
            'kg_mom_pct': None if pct(tot_kg, tot_kg_prev) is None else round(pct(tot_kg, tot_kg_prev), 1),
            'kg_yoy_pct': None if pct(tot_kg, tot_kg_yoy) is None else round(pct(tot_kg, tot_kg_yoy), 1),
            'ytd': round(ytd), 'ytd_prev': round(ytd_prev),
            'ytd_pct': None if pct(ytd, ytd_prev) is None else round(pct(ytd, ytd_prev), 1),
            'ytd_kg': round(ytd_kg), 'ytd_kg_prev': round(ytd_kg_prev),
        },
        'ex_oem': None if not has_oem else {
            'pots': round(ex_pots), 'prev': round(ex_prev), 'yoy': round(ex_yoy_v),
            'mom_pct': None if ex_mom is None else round(ex_mom, 1),
            'yoy_pct': None if ex_yoy_pct is None else round(ex_yoy_pct, 1),
        },
        'annotated': [
            {'料號': meta[k]['料號'], '品項': meta[k]['品項'],
             'status': o.get('status'), 'label': o.get('label', ''), 'note': o.get('note', '')}
            for k, o in sorted(ov.items(), key=lambda kv: meta[kv[0]]['料號'])
        ],
        'alerts': alerts, 'real_alert_count': len(real_alerts), 'core_real': core_real,
        'dormant': dormant, 'growth': growth[:cfg['top_n']],
        'newbies': newbies, 'returning': returning,
        'top_items': top_list, 'top_share': top_share,
        'categories': cats, 'same_month': same_month,
        'typical_mom': typical_mom, 'seasonal_note': seasonal_note,
        'trend': trend,
        'light': light, 'light_label': light_label, 'light_reasons': reasons,
        'config': cfg,
    }


# ────────────────────────── 輸出 ──────────────────────────
import html as _html


def esc(s):
    return _html.escape(str(s))


def fmt(n):
    try:
        return '{:,}'.format(int(round(float(n))))
    except Exception:
        return '—'


def pct_txt(v, suffix='%'):
    if v is None:
        return '—'
    return ('+' if v >= 0 else '') + ('%.1f' % v) + suffix


def pct_span(v, label=''):
    """業務量：上升為好（綠），下降為壞（紅）。"""
    if v is None:
        return '<span class="flat">—</span>'
    cls = 'up' if v > 0 else ('dn' if v < 0 else 'flat')
    arrow = '▲ ' if v > 0 else ('▼ ' if v < 0 else '')
    return '<span class="%s">%s%s%s</span>' % (cls, arrow, pct_txt(v), (' ' + label) if label else '')


def render_narrative(text):
    if not text:
        return ('<p class="muted">（本月總評尚未撰寫。由 Claude 依 '
                'reports/%s.json 的數據判讀後補上。）</p>')
    out = []
    for para in re.split(r'\n\s*\n', text.strip()):
        line = esc(para.strip()).replace('\n', '<br>')
        line = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', line)
        out.append('<p>' + line + '</p>')
    return '\n'.join(out)


VERDICT_BADGE = {
    'real': '<span class="tag t-red">🔴 實質接單下滑</span>',
    'partial': '<span class="tag t-amber">🟡 部分下滑</span>',
    'batch': '<span class="tag t-grey">⚪ 批量調整</span>',
    'unknown': '<span class="tag t-grey">－ 無重量資料</span>',
}


def alert_rows(R):
    if not R['alerts']:
        return '<tr><td colspan="8" class="empty">本月沒有達到警示門檻的品項 👍</td></tr>'
    rows = []
    for a in R['alerts']:
        core = '<span class="tag t-core">主力</span> ' if a['core'] else ''
        if a.get('tag'):
            core += '<span class="tag t-note">%s</span> ' % esc(a['tag'])
        rows.append(
            '<tr class="%s"><td>%s%s<div class="sub">%s</div></td>'
            '<td class="num">%s</td><td class="num">%s</td><td class="num">%s</td>'
            '<td class="num">%s</td><td class="num">%s</td><td class="num loss">-%s</td><td>%s</td></tr>' % (
                'row-real' if a['verdict'] == 'real' else '',
                core, esc(a['品項']), esc(a['料號']),
                fmt(a['pots_cur']), fmt(a['pots_prev']), fmt(a['pots_yoy']),
                pct_span(a['mom_pct']), pct_span(a['yoy_pct']),
                fmt(a['loss']), VERDICT_BADGE.get(a['verdict'], '')))
    return '\n'.join(rows)


def dormant_rows(R):
    if not R['dormant']:
        return '<tr><td colspan="4" class="empty">本月沒有斷單品項 👍</td></tr>'
    return '\n'.join(
        '<tr><td>%s%s<div class="sub">%s</div></td><td class="num">%s</td>'
        '<td class="num">%s</td><td class="num">%s</td></tr>' % (
            ('<span class="tag t-core">主力</span> ' if d['core'] else '')
            + ('<span class="tag t-note">%s</span> ' % esc(d['tag']) if d.get('tag') else ''),
            esc(d['品項']), esc(d['料號']),
            d['active_months'], fmt(d['avg_pots']), fmt(d['last_pots']))
        for d in R['dormant'])


def growth_rows(R):
    if not R['growth']:
        return '<tr><td colspan="6" class="empty">本月沒有明顯成長的品項</td></tr>'
    return '\n'.join(
        '<tr><td>%s%s<div class="sub">%s</div></td><td class="num">%s</td><td class="num">%s</td>'
        '<td class="num">%s</td><td class="num gain">+%s</td><td class="num">%s</td></tr>' % (
            '<span class="tag t-note">%s</span> ' % esc(g['tag']) if g.get('tag') else '',
            esc(g['品項']), esc(g['料號']), fmt(g['pots_cur']), fmt(g['pots_prev']),
            fmt(g['pots_yoy']), fmt(g['gain']), pct_span(g['mom_pct']))
        for g in R['growth'])


def newret_rows(R):
    rows = []
    for n in R['newbies']:
        rows.append('<tr><td><span class="tag t-blue">新品</span> %s<div class="sub">%s</div></td>'
                    '<td class="num">%s</td><td>—</td></tr>' % (esc(n['品項']), esc(n['料號']), fmt(n['pots'])))
    for n in R['returning']:
        rows.append('<tr><td><span class="tag t-green">回流</span> %s<div class="sub">%s</div></td>'
                    '<td class="num">%s</td><td>上次生產 %s</td></tr>' % (
                        esc(n['品項']), esc(n['料號']), fmt(n['pots']), esc(n['last_seen'])))
    return '\n'.join(rows) if rows else '<tr><td colspan="3" class="empty">本月無新品或回流品項</td></tr>'


def top_rows(R):
    return '\n'.join(
        '<tr><td>%s<div class="sub">%s</div></td><td class="num">%s</td>'
        '<td class="num">%s%%</td><td class="num">%s%%</td></tr>' % (
            esc(t['品項']), esc(t['料號']), fmt(t['pots']), t['share'], t['cum_share'])
        for t in R['top_items'])


TPL = r"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>生產鍋數月報 __MONTH__</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{--bg:#f6f7f9;--surface:#fff;--border:#e2e8f0;--text:#0f172a;--t2:#334155;--t3:#64748b;--t4:#94a3b8;
--blue:#1d4ed8;--amber:#b45309;--teal:#0f766e;--red:#dc2626;--green:#16a34a;--r:12px;
--sh:0 1px 3px rgba(15,23,42,.06);--mono:'DM Mono',ui-monospace,SFMono-Regular,Menlo,monospace;
--font:'Noto Sans TC',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:var(--font);line-height:1.6;}
.wrap{max-width:1180px;margin:0 auto;padding:26px 22px 60px;}
header.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;margin-bottom:20px;}
h1{font-size:23px;margin:0 0 4px;letter-spacing:-.01em;}
.sub{font-size:12px;color:var(--t3);}
.light{display:inline-flex;align-items:center;gap:8px;padding:9px 16px;border-radius:999px;font-weight:700;font-size:14px;}
.light.red{background:#fee2e2;color:#991b1b;} .light.yellow{background:#fef3c7;color:#92400e;} .light.green{background:#dcfce7;color:#166534;}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;box-shadow:var(--sh);margin-bottom:18px;}
.card h2{font-size:16px;margin:0 0 3px;} .card .csub{font-size:11.5px;color:var(--t3);margin-bottom:14px;}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:15px 18px;box-shadow:var(--sh);position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:linear-gradient(180deg,#1d4ed8,#6366f1);}
.kpi .lab{font-size:12px;color:var(--t3);margin-bottom:6px;}
.kpi .val{font-size:26px;font-weight:700;font-family:var(--mono);letter-spacing:-.02em;line-height:1.1;}
.kpi .dlt{font-size:11.5px;margin-top:5px;font-family:var(--mono);color:var(--t3);}
.up{color:var(--green);} .dn{color:var(--red);} .flat{color:var(--t4);}
.note{padding:11px 15px;border-radius:9px;font-size:13px;margin-bottom:18px;}
.note.warn{background:#fef3c7;border:1px solid #f6d68a;color:#92400e;}
.note.info{background:#eff6ff;border:1px solid #bfdbfe;color:#1e40af;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th,td{padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top;}
th{background:#f8fafc;font-size:11.5px;color:var(--t3);font-weight:600;white-space:nowrap;}
td.num{text-align:right;font-family:var(--mono);white-space:nowrap;}
th.num{text-align:right;}
td .sub{font-size:10.5px;color:var(--t4);font-family:var(--mono);}
td.loss{color:var(--red);font-weight:600;} td.gain{color:var(--green);font-weight:600;}
tr.row-real{background:#fff7f7;}
td.empty{text-align:center;color:var(--t4);padding:20px;}
.tag{display:inline-block;font-size:10.5px;padding:2px 7px;border-radius:20px;white-space:nowrap;}
.t-red{background:#fee2e2;color:#991b1b;} .t-amber{background:#fef3c7;color:#92400e;}
.t-grey{background:#f1f5f9;color:#64748b;} .t-core{background:#e0e7ff;color:#3730a3;font-weight:700;}
.t-blue{background:#dbeafe;color:#1e40af;} .t-green{background:#dcfce7;color:#166534;}
.t-note{background:#ede9fe;color:#5b21b6;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.chart{position:relative;height:300px;}
.chart.tall{height:360px;}
.tbl-wrap{overflow-x:auto;}
.muted{color:var(--t4);font-size:13px;}
.narrative p{margin:0 0 11px;font-size:14px;color:var(--t2);}
footer{margin-top:26px;font-size:11.5px;color:var(--t4);line-height:1.8;}
@media(max-width:860px){.kpis{grid-template-columns:1fr 1fr;}.g2{grid-template-columns:1fr;}.wrap{padding:16px 13px 44px;}}
@media print{body{background:#fff}.card{break-inside:avoid}}
</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <div>
    <h1>生產鍋數月報 · __MONTH__</h1>
    <div class="sub">業務量／接單趨勢分析　｜　與 __PREV__（上月）、__YOY__（去年同月）比較　｜　產生於 __GEN__</div>
  </div>
  <div class="light __LIGHT__">__LIGHT_ICON__ __LIGHT_LABEL__</div>
</header>

__WARN__

<div class="kpis">
  <div class="kpi"><div class="lab">本月總鍋數</div><div class="val">__POTS__</div>
    <div class="dlt">環比 __POTS_MOM__ ｜ 同比 __POTS_YOY__</div></div>
  <div class="kpi"><div class="lab">本月總重量 (kg)</div><div class="val">__KG__</div>
    <div class="dlt">環比 __KG_MOM__ ｜ 同比 __KG_YOY__</div></div>
  <div class="kpi"><div class="lab">生產品項數</div><div class="val">__ITEMS__</div>
    <div class="dlt">前 __TOPN__ 大占 __TOPSHARE__%</div></div>
  <div class="kpi"><div class="lab">接單下滑警示</div><div class="val">__ALERTS__</div>
    <div class="dlt">斷單 __DORMANT__ 項</div></div>
</div>

<div class="card">
  <h2>本月總評</h2>
  <div class="csub">綜合判讀：__LIGHT_LABEL__ — __REASONS__</div>
  <div class="narrative">__NARRATIVE__</div>
  __EXOEM__
  __SEASON__
</div>

<div class="card">
  <h2>整體業務量趨勢</h2>
  <div class="csub">近 __TRENDN__ 個月：柱狀為鍋數、折線為半成品重量(kg)。兩者背離時代表批量結構改變，而非單純接單增減。</div>
  <div class="chart tall"><canvas id="trendChart"></canvas></div>
</div>

<div class="g2">
  <div class="card"><h2>本月 vs 上月 vs 去年同月</h2><div class="csub">鍋數與重量雙軌對照</div>
    <div class="chart"><canvas id="cmpChart"></canvas></div></div>
  <div class="card"><h2>季節性對照</h2><div class="csub">歷年同月份（__MONTHNUM__月）總鍋數，用以區分季節性淡季與實質衰退</div>
    <div class="chart"><canvas id="seasonChart"></canvas></div></div>
</div>

<div class="card">
  <h2>接單下滑警示 — 流失鍋數 Top __TOPN__</h2>
  <div class="csub">與上月／去年同月相比降幅達 __DECLINE__% 且基準量 ≥ __SCALE__ 鍋者，依流失鍋數由大到小排序</div>
  <div class="chart tall"><canvas id="declineChart"></canvas></div>
</div>

<div class="card">
  <h2>接單下滑警示明細</h2>
  <div class="csub">「判讀」欄以半成品重量交叉驗證：鍋數下降但重量持平，代表批量調整而非接單下滑</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">本月</th><th class="num">上月</th><th class="num">去年同月</th>
      <th class="num">環比</th><th class="num">同比</th><th class="num">流失鍋數</th><th>判讀</th></tr></thead>
    <tbody>__ALERT_ROWS__</tbody>
  </table></div>
</div>

<div class="card">
  <h2>斷單警示</h2>
  <div class="csub">前 __LOOKBACK__ 個月中有 ≥__MINACTIVE__ 個月正常生產，但本月完全沒有生產</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">前期活躍月數</th><th class="num">平均月鍋數</th><th class="num">上月鍋數</th></tr></thead>
    <tbody>__DORMANT_ROWS__</tbody>
  </table></div>
</div>

<div class="g2">
  <div class="card"><h2>成長品項 Top __TOPN__</h2><div class="csub">依鍋數增加量排序</div>
    <div class="chart"><canvas id="growthChart"></canvas></div></div>
  <div class="card"><h2>主力集中度（Pareto）</h2><div class="csub">前 __TOPN__ 大品項鍋數與累積佔比</div>
    <div class="chart"><canvas id="paretoChart"></canvas></div></div>
</div>

<div class="card">
  <h2>成長品項明細</h2>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">本月</th><th class="num">上月</th><th class="num">去年同月</th>
      <th class="num">增加鍋數</th><th class="num">環比</th></tr></thead>
    <tbody>__GROWTH_ROWS__</tbody>
  </table></div>
</div>

<div class="g2">
  <div class="card"><h2>類別表現</h2><div class="csub">果醬餡類 vs 花生類</div>
    <div class="chart"><canvas id="catChart"></canvas></div></div>
  <div class="card"><h2>新品 / 回流品項</h2><div class="csub">本月首次生產，或中斷 3 個月以上後恢復</div>
    <div class="tbl-wrap"><table><thead><tr><th>品項</th><th class="num">本月鍋數</th><th>備註</th></tr></thead>
      <tbody>__NEWRET_ROWS__</tbody></table></div></div>
</div>

<div class="card">
  <h2>本月主力品項</h2>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">鍋數</th><th class="num">佔比</th><th class="num">累積佔比</th></tr></thead>
    <tbody>__TOP_ROWS__</tbody>
  </table></div>
</div>

<footer>
  資料來源：<code>data/latest.xlsx</code> →「資料總表」　｜　分析邏輯與門檻：<code>docs/monthly-report-logic.md</code><br>
  本報告聚焦業務量（鍋數為主、半成品重量交叉驗證）；製成率／品質分析由「產品工時及製成率統計系統」負責，不在本報告範圍。<br>
  警示門檻：降幅 ≥__DECLINE__%、基準量 ≥__SCALE__ 鍋　｜　斷單：前 __LOOKBACK__ 個月 ≥__MINACTIVE__ 月有生產而本月為 0
  __ANNOTATED__
</footer>
</div>

<script>
const R = __DATA__;
const C = {blue:'#1d4ed8', amber:'#b45309', teal:'#0f766e', red:'#dc2626', green:'#16a34a', grey:'#94a3b8'};
const grid = {color:'#eef2f7'}, tick = {color:'#64748b', font:{size:11}};
const baseOpts = (extra) => Object.assign({
  responsive:true, maintainAspectRatio:false,
  interaction:{mode:'index', intersect:false},
  plugins:{legend:{labels:{color:'#334155', font:{size:12}, usePointStyle:true, pointStyle:'circle'}}}
}, extra || {});
const hOpts = () => ({
  responsive:true, maintainAspectRatio:false, indexAxis:'y',
  interaction:{mode:'index', intersect:false, axis:'y'},
  plugins:{legend:{labels:{color:'#334155', font:{size:12}, usePointStyle:true, pointStyle:'circle'}}},
  scales:{x:{grid:grid, ticks:tick, beginAtZero:true}, y:{grid:{display:false}, ticks:{color:'#0f172a', font:{size:11}}}}
});
const shortName = (s) => s.length > 16 ? s.slice(0,16)+'…' : s;

// 1. 趨勢
new Chart(document.getElementById('trendChart'), {
  data:{labels:R.trend.map(t=>t.label), datasets:[
    {type:'bar', label:'鍋數', data:R.trend.map(t=>t.pots), backgroundColor:C.blue+'cc', borderRadius:4, yAxisID:'y'},
    {type:'line', label:'重量 (kg)', data:R.trend.map(t=>t.kg), borderColor:C.teal, backgroundColor:C.teal,
     borderWidth:2, tension:.3, pointRadius:2, yAxisID:'y1'}
  ]},
  options: baseOpts({scales:{
    x:{grid:grid, ticks:Object.assign({maxRotation:60, autoSkip:false}, tick)},
    y:{position:'left', grid:grid, ticks:tick, title:{display:true, text:'鍋數', color:'#64748b', font:{size:11}}},
    y1:{position:'right', grid:{display:false}, ticks:tick, title:{display:true, text:'kg', color:'#64748b', font:{size:11}}}
  }})
});

// 2. 三期比較
new Chart(document.getElementById('cmpChart'), {
  type:'bar',
  data:{labels:['本月 '+R.month, '上月 '+R.prev_label, '去年同月 '+R.yoy_label], datasets:[
    {label:'鍋數', data:[R.totals.pots, R.totals.pots_prev, R.totals.pots_yoy], backgroundColor:C.blue+'cc', borderRadius:5, yAxisID:'y'},
    {label:'重量 (kg)', data:[R.totals.kg, R.totals.kg_prev, R.totals.kg_yoy], backgroundColor:C.teal+'99', borderRadius:5, yAxisID:'y1'}
  ]},
  options: baseOpts({scales:{
    x:{grid:{display:false}, ticks:tick},
    y:{position:'left', grid:grid, ticks:tick, title:{display:true, text:'鍋數', color:'#64748b', font:{size:11}}},
    y1:{position:'right', grid:{display:false}, ticks:tick, title:{display:true, text:'kg', color:'#64748b', font:{size:11}}}
  }})
});

// 3. 季節性
new Chart(document.getElementById('seasonChart'), {
  type:'bar',
  data:{labels:R.same_month.map(s=>s.year+'年'), datasets:[
    {label:R.mon+'月總鍋數', data:R.same_month.map(s=>s.pots),
     backgroundColor:R.same_month.map(s=>s.year===R.year?C.blue+'dd':C.grey+'99'), borderRadius:5}
  ]},
  options: baseOpts({plugins:{legend:{display:false}},
    scales:{x:{grid:{display:false}, ticks:tick}, y:{grid:grid, ticks:tick, beginAtZero:true}}})
});

// 4. 衰退 Top N
const dec = R.alerts.slice(0, R.config.top_n);
new Chart(document.getElementById('declineChart'), {
  type:'bar',
  data:{labels:dec.map(a=>shortName(a['品項'])), datasets:[
    {label:'較上月流失鍋數', data:dec.map(a=>Math.max(0, a.pots_prev-a.pots_cur)), backgroundColor:C.red+'bb', borderRadius:4},
    {label:'較去年同月流失鍋數', data:dec.map(a=>Math.max(0, a.pots_yoy-a.pots_cur)), backgroundColor:C.amber+'99', borderRadius:4}
  ]},
  options: hOpts()
});

// 5. 成長 Top N
new Chart(document.getElementById('growthChart'), {
  type:'bar',
  data:{labels:R.growth.map(g=>shortName(g['品項'])), datasets:[
    {label:'增加鍋數', data:R.growth.map(g=>g.gain), backgroundColor:C.green+'bb', borderRadius:4}
  ]},
  options: hOpts()
});

// 6. Pareto
new Chart(document.getElementById('paretoChart'), {
  data:{labels:R.top_items.map(t=>shortName(t['品項'])), datasets:[
    {type:'bar', label:'鍋數', data:R.top_items.map(t=>t.pots), backgroundColor:C.blue+'bb', borderRadius:4, yAxisID:'y'},
    {type:'line', label:'累積佔比 %', data:R.top_items.map(t=>t.cum_share), borderColor:C.amber,
     backgroundColor:C.amber, borderWidth:2, tension:.25, pointRadius:3, yAxisID:'y1'}
  ]},
  options: baseOpts({scales:{
    x:{grid:{display:false}, ticks:Object.assign({maxRotation:60, autoSkip:false}, tick)},
    y:{position:'left', grid:grid, ticks:tick},
    y1:{position:'right', grid:{display:false}, ticks:tick, min:0, max:100}
  }})
});

// 7. 類別
const cats = Object.keys(R.categories);
new Chart(document.getElementById('catChart'), {
  type:'bar',
  data:{labels:cats, datasets:[
    {label:'本月', data:cats.map(c=>R.categories[c].pots), backgroundColor:C.blue+'cc', borderRadius:5},
    {label:'上月', data:cats.map(c=>R.categories[c].prev), backgroundColor:C.grey+'99', borderRadius:5},
    {label:'去年同月', data:cats.map(c=>R.categories[c].yoy), backgroundColor:C.amber+'99', borderRadius:5}
  ]},
  options: baseOpts({scales:{x:{grid:{display:false}, ticks:tick}, y:{grid:grid, ticks:tick, beginAtZero:true}}})
});
</script>
</body>
</html>
"""

LIGHT_ICON = {'red': '🔴', 'yellow': '🟡', 'green': '🟢'}


def build_html(R, narrative):
    t = R['totals']
    cfg = R['config']
    warn = ''
    if R['incomplete']:
        warn = ('<div class="note warn">⚠️ <strong>%s 的資料可能尚未補齊</strong>：本月僅 %d 筆紀錄，'
                '約為其他月份中位數 %d 的 %d%%。報告數字可能低估，請確認資料是否完整後重新產生。</div>' % (
                    R['month'], R['record_count'], R['median_count'],
                    round(R['record_count'] / R['median_count'] * 100) if R['median_count'] else 0))
    exo = ''
    x = R.get('ex_oem')
    if x:
        exo = ('<div class="note info" style="margin:14px 0 0">🏭 <strong>排除季節性代工後：</strong>'
               '%s 鍋（環比 %s ／ 同比 %s）—— 代工品項量體大且間歇，扣除後較能反映本業的真實表現。</div>'
               % (fmt(x['pots']), pct_span(x['mom_pct']), pct_span(x['yoy_pct'])))

    ann = ''
    if R.get('annotated'):
        STATUS_EFFECT = {'discontinued': '不列入警示與燈號',
                         'seasonal_oem': '不列主力、不觸發燈號、另計排除後總量'}
        parts = ['<strong>%s</strong> %s（%s）%s' % (
            esc(a['label']), esc(a['品項']), esc(a['料號']),
            STATUS_EFFECT.get(a['status'], '')) for a in R['annotated']]
        ann = '<br>品項註記：' + '　｜　'.join(parts)

    season = ''
    if R['seasonal_note']:
        season = '<div class="note info" style="margin:14px 0 0">📅 <strong>季節性判讀：</strong>%s</div>' % esc(R['seasonal_note'])

    repl = {
        '__MONTH__': R['month'], '__PREV__': R['prev_label'], '__YOY__': R['yoy_label'],
        '__GEN__': R['generated_at'], '__MONTHNUM__': str(R['mon']),
        '__LIGHT__': R['light'], '__LIGHT_LABEL__': R['light_label'],
        '__LIGHT_ICON__': LIGHT_ICON.get(R['light'], ''),
        '__REASONS__': esc('；'.join(R['light_reasons'])),
        '__WARN__': warn, '__SEASON__': season,
        '__EXOEM__': exo, '__ANNOTATED__': ann,
        '__POTS__': fmt(t['pots']), '__KG__': fmt(t['kg']), '__ITEMS__': str(t['items']),
        '__POTS_MOM__': pct_span(t['mom_pct']), '__POTS_YOY__': pct_span(t['yoy_pct']),
        '__KG_MOM__': pct_span(t['kg_mom_pct']), '__KG_YOY__': pct_span(t['kg_yoy_pct']),
        '__ALERTS__': str(R['real_alert_count']), '__DORMANT__': str(len(R['dormant'])),
        '__TOPSHARE__': str(R['top_share']), '__TOPN__': str(cfg['top_n']),
        '__DECLINE__': str(cfg['decline_pct']), '__SCALE__': str(cfg['min_scale_pots']),
        '__LOOKBACK__': str(cfg['dormant_lookback']), '__MINACTIVE__': str(cfg['dormant_min_active']),
        '__TRENDN__': str(len(R['trend'])),
        '__NARRATIVE__': render_narrative(narrative) if narrative else render_narrative(None) % R['month'],
        '__ALERT_ROWS__': alert_rows(R), '__DORMANT_ROWS__': dormant_rows(R),
        '__GROWTH_ROWS__': growth_rows(R), '__NEWRET_ROWS__': newret_rows(R),
        '__TOP_ROWS__': top_rows(R),
        '__DATA__': json.dumps(R, ensure_ascii=False).replace('</', '<\\/'),
    }
    out = TPL
    for k, v in repl.items():
        out = out.replace(k, v)
    return out


def build_index():
    entries = []
    for fn in sorted(os.listdir(REPORTS), reverse=True):
        m = re.match(r'^(\d{4})-(\d{2})\.json$', fn)
        if not m:
            continue
        try:
            with open(os.path.join(REPORTS, fn), encoding='utf-8') as f:
                R = json.load(f)
        except Exception:
            continue
        entries.append(R)
    rows = []
    for R in entries:
        t = R['totals']
        rows.append(
            '<a class="row" href="%s-%s.html"><div class="m">%s</div>'
            '<div class="light %s">%s %s</div>'
            '<div class="n">%s 鍋　環比 %s　同比 %s</div>'
            '<div class="a">警示 %d ｜ 斷單 %d</div></a>' % (
                R['month'][:4], R['month'][5:7], esc(R['month']),
                R['light'], LIGHT_ICON.get(R['light'], ''), esc(R['light_label']),
                fmt(t['pots']), pct_span(t['mom_pct']), pct_span(t['yoy_pct']),
                R['real_alert_count'], len(R['dormant'])))
    body = '\n'.join(rows) if rows else '<p class="muted">尚未產生任何月報。</p>'
    html_out = """<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>生產鍋數月報 · 索引</title>
<style>
body{margin:0;background:#f6f7f9;color:#0f172a;font-family:'Noto Sans TC',-apple-system,'Segoe UI',sans-serif;}
.wrap{max-width:820px;margin:0 auto;padding:34px 20px 60px;}
h1{font-size:22px;margin:0 0 4px;} .sub{font-size:12.5px;color:#64748b;margin-bottom:22px;}
.row{display:grid;grid-template-columns:88px 110px 1fr auto;gap:14px;align-items:center;
 background:#fff;border:1px solid #e2e8f0;border-radius:11px;padding:14px 17px;margin-bottom:10px;
 text-decoration:none;color:inherit;box-shadow:0 1px 3px rgba(15,23,42,.05);}
.row:hover{border-color:#1d4ed8;}
.m{font-family:ui-monospace,monospace;font-weight:700;font-size:15px;}
.light{font-size:11.5px;padding:3px 9px;border-radius:20px;text-align:center;white-space:nowrap;}
.light.red{background:#fee2e2;color:#991b1b;}.light.yellow{background:#fef3c7;color:#92400e;}.light.green{background:#dcfce7;color:#166534;}
.n{font-size:12.5px;color:#334155;font-family:ui-monospace,monospace;}
.a{font-size:11.5px;color:#64748b;white-space:nowrap;}
.up{color:#16a34a;}.dn{color:#dc2626;}.flat{color:#94a3b8;}
.muted{color:#94a3b8;font-size:13px;}
a.back{display:inline-block;margin-bottom:18px;font-size:12.5px;color:#1d4ed8;text-decoration:none;}
@media(max-width:640px){.row{grid-template-columns:1fr;gap:5px;}}
</style></head><body><div class="wrap">
<a class="back" href="../">← 回到儀表板</a>
<h1>生產鍋數月報</h1>
<div class="sub">依業務量（鍋數／重量）分析接單趨勢與衰退警示。點擊月份查看完整報告。</div>
__ROWS__
</div></body></html>"""
    return html_out.replace('__ROWS__', body)


def main():
    ap = argparse.ArgumentParser(description='產生月度業務分析報告')
    ap.add_argument('--month', required=True, help='目標月份，格式 YYYY-MM，例如 2026-08')
    ap.add_argument('--narrative', help='總評文字檔（純文字／段落以空行分隔，支援 **粗體**）')
    ap.add_argument('--xlsx', default=XLSX, help='資料來源 xlsx（預設 data/latest.xlsx）')
    ap.add_argument('--allow-incomplete', action='store_true', help='即使該月未補齊仍產生報告')
    args = ap.parse_args()

    m = re.match(r'^(\d{4})-(\d{1,2})$', args.month.strip())
    if not m:
        raise SystemExit('月份格式錯誤，請用 YYYY-MM，例如 2026-08')
    Y, M = int(m.group(1)), int(m.group(2))
    if not 1 <= M <= 12:
        raise SystemExit('月份必須介於 1~12')

    with open(CONFIG, encoding='utf-8') as f:
        cfg = json.load(f)

    if not os.path.exists(args.xlsx):
        raise SystemExit('找不到資料檔 %s' % args.xlsx)
    rows = read_sheet(args.xlsx, SHEET)
    pots, kg, meta, month_counts = aggregate(rows)

    if (Y, M) not in month_counts:
        have = sorted(month_counts)
        raise SystemExit('資料中沒有 %s 的紀錄。目前資料涵蓋 %s ~ %s。' % (
            mlabel(Y, M), mlabel(*have[0]), mlabel(*have[-1])))

    R = compute(pots, kg, meta, month_counts, Y, M, cfg)
    if R['incomplete'] and not args.allow_incomplete:
        raise SystemExit(
            '%s 的資料可能尚未補齊（僅 %d 筆，約為其他月份中位數 %d 的 %d%%）。\n'
            '請確認資料完整後再產生；若確定要產出，加上 --allow-incomplete。' % (
                R['month'], R['record_count'], R['median_count'],
                round(R['record_count'] / R['median_count'] * 100) if R['median_count'] else 0))

    narrative = None
    if args.narrative:
        with open(args.narrative, encoding='utf-8') as f:
            narrative = f.read()

    os.makedirs(REPORTS, exist_ok=True)
    stem = '%04d-%02d' % (Y, M)
    with open(os.path.join(REPORTS, stem + '.json'), 'w', encoding='utf-8') as f:
        json.dump(R, f, ensure_ascii=False, indent=1)
    with open(os.path.join(REPORTS, stem + '.html'), 'w', encoding='utf-8') as f:
        f.write(build_html(R, narrative))
    with open(os.path.join(REPORTS, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(build_index())

    t = R['totals']
    print('✓ 已產生 reports/%s.html' % stem)
    print('  總鍋數 %s（環比 %s／同比 %s）｜ 重量 %s kg ｜ 品項 %d' % (
        fmt(t['pots']), pct_txt(t['mom_pct']), pct_txt(t['yoy_pct']), fmt(t['kg']), t['items']))
    print('  燈號 %s %s ｜ 警示 %d 項 ｜ 斷單 %d 項' % (
        LIGHT_ICON.get(R['light'], ''), R['light_label'], R['real_alert_count'], len(R['dormant'])))


if __name__ == '__main__':
    main()
