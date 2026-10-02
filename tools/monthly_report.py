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
import calendar
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
C_SEMI, C_FG = 7, 15   # 半成品數(實際)、成品數(實際)：完整性檢查用（空白＝生管尚未填寫；0 為有效值）


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
    """回傳 (recs, meta, month_counts, last_date)。

    recs 為逐筆紀錄 (key, (y, m), day, pots, kg)，已排除「刪除」列；
    月彙總由 roll() 產生，期中預覽可只取每月前 N 天做同期比較。
    """
    recs, meta = [], {}
    month_counts = {}
    last_date = None
    for r in rows[1:]:
        ymd = parse_ymd(r.get(C_DATE))
        if not ymd:
            continue
        y, m, d = ymd
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
        if last_date is None or ymd > last_date:
            last_date = ymd
        recs.append((key, ym, d, to_num(r.get(C_POTS)), to_num(r.get(C_KG))))
    return recs, meta, month_counts, last_date


def blank_rows(rows, Y, M, cap_day=None, skip_codes=()):
    """本月有鍋數、但「半成品數(實際)」或「成品數(實際)」為空白的紀錄（生管尚未填寫）。
    填 0 視為已填（例如重工、不出貨）；回傳 [(日, 料號, 品項, 缺漏欄位)]。"""
    blank = lambda v: v is None or str(v).strip() == ''
    out = []
    for r in rows[1:]:
        ymd = parse_ymd(r.get(C_DATE))
        if not ymd or (ymd[0], ymd[1]) != (Y, M) or (cap_day and ymd[2] > cap_day):
            continue
        name = r.get(C_NAME)
        if name is None or str(name).strip() in ('#N/A', '') or r.get(C_DELETED):
            continue
        p = to_num(r.get(C_POTS))
        if not p or p <= 0 or str(r.get(C_CODE) or '').strip() in skip_codes:
            continue
        miss = [lbl for c, lbl in ((C_SEMI, '半成品數'), (C_FG, '成品數')) if blank(r.get(c))]
        if miss:
            out.append((ymd[2], str(r.get(C_CODE) or ''), str(name), '／'.join(miss)))
    return out


def roll(recs, cap_day=None):
    """逐筆紀錄 → 月彙總 (pots, kg, nokg)。cap_day 有值時每個月只計入 1~cap_day 日（同期比較用）。
    nokg＝有鍋數但重量未填（空白或 0）的鍋數，用來區分「重量尚未補登」與真正的異常。"""
    pots, kg, nokg = {}, {}, {}
    for key, ym, d, pv, kv in recs:
        if cap_day is not None and d > cap_day:
            continue
        if pv is not None:
            pots.setdefault(key, {})[ym] = pots.setdefault(key, {}).get(ym, 0.0) + pv
            if pv > 0 and not kv:
                nokg.setdefault(key, {})[ym] = nokg.setdefault(key, {}).get(ym, 0.0) + pv
        if kv is not None:
            kg.setdefault(key, {})[ym] = kg.setdefault(key, {}).get(ym, 0.0) + kv
    return pots, kg, nokg


# ────────────────────────── 分析 ──────────────────────────
def r1(v):
    return None if v is None else round(v, 1)


def compute(full, cmp, meta, month_counts, Y, M, cfg, preview=None, with_tracking=True):
    """full=(pots, kg) 完整月彙總（歷史／主力排名／斷單回溯用）；
    cmp=(pots, kg) 比較用彙總——正式月報與 full 相同，期中預覽則為「每月前 N 天」的同期資料。"""
    pots, kg = full[0], full[1]
    pots_c, kg_c, nokg_c = cmp
    Pf = lambda k, ym: pots.get(k, {}).get(ym, 0.0)
    Kf = lambda k, ym: kg.get(k, {}).get(ym, 0.0)
    P = lambda k, ym: pots_c.get(k, {}).get(ym, 0.0)
    K = lambda k, ym: kg_c.get(k, {}).get(ym, 0.0)
    NK = lambda k, ym: nokg_c.get(k, {}).get(ym, 0.0)   # 重量未填的鍋數

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

    def month_complete(ym):
        """歷史月份是否資料完整（筆數 ≥ 中位數 50%）；季節性對照用。"""
        return bool(median) and month_counts.get(ym, 0) >= 0.5 * median

    # ── 整體（期中預覽時為同期數字）──
    tot_pots = sum(P(k, cur) for k in keys)
    tot_pots_prev = sum(P(k, prev) for k in keys)
    tot_pots_yoy = sum(P(k, yoy) for k in keys)
    tot_kg = sum(K(k, cur) for k in keys)
    tot_kg_prev = sum(K(k, prev) for k in keys)
    tot_kg_yoy = sum(K(k, yoy) for k in keys)
    active_items = [k for k in keys if P(k, cur) > 0]
    # 重量常於事後補登：本月有重量的鍋數占比過低時，視為「重量尚未填入」
    nokg_pots = sum(NK(k, cur) for k in keys)
    kg_cov = ((tot_pots - nokg_pots) / tot_pots * 100) if tot_pots else 0.0
    kg_pending = kg_cov < float(cfg.get('anomaly_kg_coverage_pct', 50))

    # ── 主力品項：以「前 12 個月累計鍋數」排名（不受本月衰退影響）──
    base_window = [month_add(Y, M, -i) for i in range(1, 13)]
    baseline = {k: sum(Pf(k, ym) for ym in base_window) for k in keys}
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
    core_ranked = sorted(rank_pool, key=lambda k: -baseline[k])[:cfg['top_n']]
    core_keys = set(core_ranked)
    pool_total = sum(baseline[k] for k in rank_pool)

    dec = float(cfg['decline_pct'])
    scale = float(cfg['min_scale_pots'])
    kg_flat = float(cfg['kg_flat_pct'])

    # ── 排除季節性代工後的總量（本業表現；燈號與 B 級門檻以此為準）──
    has_oem = bool(oem_keys) and any(
        P(k, cur) > 0 or P(k, prev) > 0 or P(k, yoy) > 0 for k in oem_keys)
    ex_pots = sum(P(k, cur) for k in keys if k not in oem_keys)
    ex_prev = sum(P(k, prev) for k in keys if k not in oem_keys)
    ex_yoy_v = sum(P(k, yoy) for k in keys if k not in oem_keys)
    ex_mom = pct(ex_pots, ex_prev)
    ex_yoy_pct = pct(ex_pots, ex_yoy_v)
    mom_pct_total = pct(tot_pots, tot_pots_prev)
    yoy_pct_total = pct(tot_pots, tot_pots_yoy)
    judge_total = ex_pots if has_oem else tot_pots
    judge_mom = ex_mom if has_oem else mom_pct_total
    judge_yoy = ex_yoy_pct if has_oem else yoy_pct_total

    # ── 接單下滑警示 ──
    # 分級只看鍋數；重量變化僅作參考（製程條件改變也會讓重量變動），不再用來降級或排除。
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
        if loss_yoy > loss_mom:
            basis, loss, p_pct = 'YoY', loss_yoy, yoy_pct
            kg_base, kg_pct_v = ky, pct(kc, ky)
        else:
            basis, loss, p_pct = 'MoM', loss_mom, mom_pct
            kg_base, kg_pct_v = kp, pct(kc, kp)

        if kg_pending or (preview and NK(k, cur) > 0):
            verdict, vlabel = 'pending', '重量尚未填入'
        elif kg_base <= 0 or kg_pct_v is None:
            verdict, vlabel = 'unknown', '無重量資料'
        elif kg_pct_v <= -dec:
            verdict, vlabel = 'real', '重量同步下降'
        elif kg_pct_v <= -kg_flat:
            verdict, vlabel = 'partial', '重量小幅下降'
        else:
            verdict, vlabel = 'batch', '重量持平或增加'

        alerts.append({
            'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'], '類別': meta[k]['類別'],
            'pots_cur': round(pc), 'pots_prev': round(pp), 'pots_yoy': round(py),
            'kg_cur': round(kc), 'kg_prev': round(kp), 'kg_yoy': round(ky),
            'mom_pct': r1(mom_pct), 'yoy_pct': r1(yoy_pct), 'kg_pct': r1(kg_pct_v),
            'basis': basis, 'loss': round(loss), 'pots_pct': r1(p_pct),
            'verdict': verdict, 'verdict_label': vlabel,
            'core': k in core_keys,
            'tag': ov[k]['label'] if k in ov else None,
            'oem': k in oem_keys,
        })

    # ── 斷單（前 N 個月常態生產、本月掛零）──
    # 期中預覽時本月尚未結束，同樣條件的品項改稱「截至目前尚未生產」，不算斷單、不觸發燈號。
    look = int(cfg['dormant_lookback'])
    need = int(cfg['dormant_min_active'])
    lost_months = int(cfg.get('dormant_lost_months', 3))
    dormant = []
    for k in keys:
        if k in dropped_keys or P(k, cur) > 0:
            continue
        window = [month_add(Y, M, -i) for i in range(1, look + 1)]
        vals = [Pf(k, ym) for ym in window]
        active = sum(1 for v in vals if v > 0)
        if active < need:
            continue
        avg = sum(vals) / max(1, active)
        if avg < scale:
            continue
        # 連續未生產月數（正式月報含本月；預覽不含尚未結束的本月）與最後生產月
        gap, last_seen, last_pots = 0, None, 0
        start = 0 if not preview else 1
        for i in range(start, 37):
            ym = month_add(Y, M, -i)
            if Pf(k, ym) > 0:
                last_seen, last_pots = ym, Pf(k, ym)
                break
            gap += 1
        dormant.append({
            'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'],
            'active_months': active, 'avg_pots': round(avg),
            'last_pots': round(vals[0]),
            'gap_months': gap,
            'last_seen': mlabel(*last_seen) if last_seen else None,
            'last_seen_pots': round(last_pots),
            'suspect_lost': (not preview) and gap >= lost_months,
            'core': k in core_keys,
            'tag': ov[k]['label'] if k in ov else None,
        })
    dormant.sort(key=lambda d: -d['avg_pots'])
    dormant_keys = {d['key'] for d in dormant}

    # 已列為斷單者不再重複列入衰退警示，讓兩份清單互斥、計數不重複
    alerts = [a for a in alerts if a['key'] not in dormant_keys]

    # ── 警示分級 ──
    # A：主力品項（不設門檻）；B：非主力且流失 ≥ 本業總鍋數 × b_grade_pct%；C：其餘（報告中收合）
    # 季節性代工品項另列「註」：量體大且間歇，放進分級會蓋過本業訊號。
    b_pct = float(cfg.get('b_grade_pct', 1))
    b_threshold = max(1, round(judge_total * b_pct / 100.0))
    for a in alerts:
        if a['oem']:
            a['grade'] = 'N'
        else:
            a['grade'] = 'A' if a['core'] else ('B' if a['loss'] >= b_threshold else 'C')
    alerts.sort(key=lambda a: ('ABNC'.index(a['grade']), -a['loss']))
    grade_counts = {g: sum(1 for a in alerts if a['grade'] == g) for g in 'ABCN'}
    a_alerts = [a for a in alerts if a['grade'] == 'A']
    core_loss = sum(a['loss'] for a in a_alerts)
    core_loss_pct = core_loss / judge_total * 100 if judge_total else None

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
            'mom_pct': r1(pct(pc, pp)), 'yoy_pct': r1(pct(pc, py)),
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
        if all(Pf(k, ym) == 0 for ym in last3):
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
    # 附錄：本月有生產的全部品項（含小量品項，避免「有生產卻在報告找不到」）
    all_items = [{'品項': meta[k]['品項'], '料號': meta[k]['料號'], '類別': meta[k]['類別'],
                  'pots': r1(P(k, cur)), 'prev': r1(P(k, prev)), 'yoy': r1(P(k, yoy)),
                  'kg': round(K(k, cur)), 'mom_pct': r1(pct(P(k, cur), P(k, prev))),
                  'tag': (ov.get(k) or {}).get('label')}
                 for k in sorted(active_items, key=lambda k: (-P(k, cur), meta[k]['料號']))]

    # ── 類別 ──
    cats = {}
    for cat in ('果醬餡類', '花生類'):
        ks = [k for k in keys if meta[k]['類別'] == cat]
        c, p, yv = (sum(P(k, cur) for k in ks), sum(P(k, prev) for k in ks), sum(P(k, yoy) for k in ks))
        cats[cat] = {
            'pots': round(c), 'prev': round(p), 'yoy': round(yv),
            'kg': round(sum(K(k, cur) for k in ks)),
            'mom_pct': r1(pct(c, p)), 'yoy_pct': r1(pct(c, yv)),
            'share': round(c / tot_pots * 100, 1) if tot_pots else 0,
        }

    # ── 季節性：同月跨年度，以及「該月轉換的常態變化」──
    # 資料不完整的歷史月份（筆數 < 中位數 50%，或整月缺漏）不列入，並在報告中註明樣本數。
    years = sorted({y for (y, _m) in month_counts})
    same_month, season_excluded = [], []
    for y in years:
        if y > Y:
            continue
        if y < Y and not month_complete((y, M)):
            if y >= years[0]:
                season_excluded.append(mlabel(y, M))
            continue
        same_month.append({'year': y, 'pots': round(sum(P(k, (y, M)) for k in keys))})
    typical_moms, season_years = [], []
    for y in years:
        if y >= Y:
            continue
        pym = month_add(y, M, -1)
        if not (month_complete((y, M)) and month_complete(pym)):
            continue
        a = sum(P(k, (y, M)) for k in keys)
        b = sum(P(k, pym) for k in keys)
        r = pct(a, b)
        if r is not None and b > 0 and a > 0:
            typical_moms.append(r)
            season_years.append(y)
    typical_mom = round(sum(typical_moms) / len(typical_moms), 1) if typical_moms else None
    seasonal_note = None
    if typical_mom is not None and mom_pct_total is not None:
        gap = mom_pct_total - typical_mom
        if mom_pct_total < 0 and gap >= -5:
            seasonal_note = '本月環比 %.1f%%，過去同月份平均為 %.1f%%，變化落在季節性常態範圍內。' % (mom_pct_total, typical_mom)
        elif gap < -5:
            seasonal_note = '本月環比 %.1f%%，明顯低於過去同月份平均的 %.1f%%，非單純季節性因素。' % (mom_pct_total, typical_mom)
        else:
            seasonal_note = '本月環比 %.1f%%，優於過去同月份平均的 %.1f%%。' % (mom_pct_total, typical_mom)
    sample_txt = '季節性參考樣本：%d 年（%s）' % (
        len(season_years), '、'.join(str(y) for y in season_years) or '無')
    if season_excluded:
        sample_txt += '；%s 資料不完整，未列入' % '、'.join(season_excluded)
    if len(season_years) < 3:
        sample_txt += '。樣本少於 3 年，僅供參考'
    seasonal_sample = sample_txt + '。'

    # ── 近 N 個月趨勢（完整月；預覽時本月為截至目前）──
    trend = []
    for i in range(int(cfg['trend_months']) - 1, -1, -1):
        ym = month_add(Y, M, -i)
        if ym not in month_counts:
            continue
        lab = mlabel(*ym)
        if preview and ym == cur:
            lab += '*'
        trend.append({'label': lab,
                      'pots': round(sum(Pf(k, ym) for k in keys)),
                      'kg': round(sum(Kf(k, ym) for k in keys))})

    # ── YTD（1月~本月）今年 vs 去年同期；本月以比較口徑計（預覽時為同期）──
    def ytd_of(year, ks):
        return (sum(Pf(k, (year, m)) for k in ks for m in range(1, M))
                + sum(P(k, (year, M)) for k in ks))
    ytd = ytd_of(Y, keys)
    ytd_prev = ytd_of(Y - 1, keys)
    ytd_kg = (sum(Kf(k, (Y, m)) for k in keys for m in range(1, M)) + sum(K(k, cur) for k in keys))
    ytd_kg_prev = (sum(Kf(k, (Y - 1, m)) for k in keys for m in range(1, M)) + sum(K(k, yoy) for k in keys))
    core_biz = [k for k in keys if k not in oem_keys]
    ytd_oem = ytd_of(Y, oem_keys) if oem_keys else 0
    ytd_oem_prev = ytd_of(Y - 1, oem_keys) if oem_keys else 0
    has_oem_ytd = bool(ytd_oem or ytd_oem_prev)
    ytd_ex = ytd_of(Y, core_biz)
    ytd_ex_prev = ytd_of(Y - 1, core_biz)
    series = {'labels': [], 'this_ex': [], 'last_ex': [], 'this_all': []}
    c_this = c_last = c_all = 0.0
    for m in range(1, M + 1):
        f = P if m == M else Pf
        c_this += sum(f(k, (Y, m)) for k in core_biz)
        c_last += sum(f(k, (Y - 1, m)) for k in core_biz)
        c_all += sum(f(k, (Y, m)) for k in keys)
        series['labels'].append('%d月' % m)
        series['this_ex'].append(round(c_this))
        series['last_ex'].append(round(c_last))
        series['this_all'].append(round(c_all))
    oem_names = '、'.join(meta[k]['品項'] for k in sorted(oem_keys) if ytd_of(Y, [k]) or ytd_of(Y - 1, [k]))
    ytd_block = {
        'has_oem': has_oem_ytd, 'oem_names': oem_names,
        'pots': round(ytd_ex), 'pots_prev': round(ytd_ex_prev), 'pct': r1(pct(ytd_ex, ytd_ex_prev)),
        'oem_pots': round(ytd_oem), 'oem_pots_prev': round(ytd_oem_prev),
        'all_pots': round(ytd), 'all_pots_prev': round(ytd_prev), 'all_pct': r1(pct(ytd, ytd_prev)),
        'series': series,
    }

    # ── 環比瀑布：上月 → 各品項增減 → 本月 ──
    wf_n = int(cfg.get('waterfall_items', 5))
    deltas = [(k, P(k, cur) - P(k, prev)) for k in keys]
    neg = sorted([d for d in deltas if d[1] < 0], key=lambda d: d[1])
    pos = sorted([d for d in deltas if d[1] > 0], key=lambda d: -d[1])
    steps = []
    for k, d in neg[:wf_n]:
        steps.append({'label': meta[k]['品項'], 'code': meta[k]['料號'], 'delta': round(d)})
    if neg[wf_n:]:
        steps.append({'label': '其他 %d 項減少' % len(neg[wf_n:]), 'code': '',
                      'delta': round(sum(d for _k, d in neg[wf_n:]))})
    if pos[wf_n:]:
        steps.append({'label': '其他 %d 項增加' % len(pos[wf_n:]), 'code': '',
                      'delta': round(sum(d for _k, d in pos[wf_n:]))})
    for k, d in reversed(pos[:wf_n]):
        steps.append({'label': meta[k]['品項'], 'code': meta[k]['料號'], 'delta': round(d)})
    waterfall = {'start': round(tot_pots_prev), 'end': round(tot_pots), 'steps': steps,
                 'dec_total': round(sum(d for _k, d in neg)), 'inc_total': round(sum(d for _k, d in pos))}

    # ── 資料異常：鍋數與重量不匹配（只點出，原因請生管單位說明）──
    # 重量常於事後補登：本月有重量的鍋數占比 < anomaly_kg_coverage_pct 時視為「重量尚未填入」，暫不檢查。
    an_min = float(cfg.get('anomaly_min_pots', 5))
    an_pct = float(cfg.get('anomaly_dev_pct', 50))
    an_hist = int(cfg.get('anomaly_min_history', 3))
    an_low = float(cfg.get('anomaly_new_low_pct', 20))
    an_high = float(cfg.get('anomaly_new_high_pct', 300))
    f_p = sum(Pf(k, ym) for k in keys for ym in base_window)
    f_k = sum(Kf(k, ym) for k in keys for ym in base_window)
    factory_kpp = f_k / f_p if f_p else None
    anomalies = []
    for k in ([] if kg_pending else keys):
        pc = P(k, cur)
        if pc < an_min:
            continue
        nk = NK(k, cur)
        filled = pc - nk          # 已填重量的鍋數；每鍋重量只用這部分計算
        kc = K(k, cur)
        hist = [Kf(k, ym) / Pf(k, ym) for ym in base_window
                if Pf(k, ym) >= an_min and Kf(k, ym) > 0]
        ratio = kc / filled if filled > 0 else 0.0
        kind, dev, med, basis = None, None, None, None
        if nk >= an_min and not preview:
            # 正式月報：月份已結束仍有批次沒有重量 → 異常（期中預覽視為尚未補登，不列）
            kind = '有鍋數、無重量（%s 鍋）' % '{:,}'.format(round(nk))
        elif filled < an_min:
            continue
        elif len(hist) >= an_hist:
            # 有足夠歷史：與品項自己的前 12 個月中位數比較
            med, basis = statistics.median(hist), '自身歷史'
            dev = (ratio / med - 1) * 100
            if abs(dev) >= an_pct:
                kind = '每鍋重量偏低' if dev < 0 else '每鍋重量偏高'
        elif factory_kpp:
            # 新品或歷史不足：與全廠平均每鍋重量比較，只抓極端值
            med, basis = factory_kpp, '全廠平均'
            dev = (ratio / med - 1) * 100
            share = ratio / med * 100
            if share < an_low or share > an_high:
                kind = '每鍋重量與全廠差異過大'
        if kind:
            anomalies.append({
                'key': k, '料號': meta[k]['料號'], '品項': meta[k]['品項'],
                'pots': round(pc), 'kg': round(kc), 'kg_per_pot': round(ratio, 1),
                'hist_median': None if med is None else round(med, 1), 'basis': basis,
                'hist_n': len(hist), 'dev_pct': r1(dev), 'kind': kind,
                'tag': ov[k]['label'] if k in ov else None,
            })
    anomalies.sort(key=lambda x: -x['pots'])
    anomaly_keys = {x['key'] for x in anomalies}

    # ── 主力品項 3 個月移動平均（只用完整月份的實際值；預覽時截至上月）──
    # 單一品項月產量受排程影響波動大，以 3 個月平均對比「去年同期 3 個月平均」判斷趨勢，
    # 同時消除季節性。僅供參考，不影響警示分級與燈號。
    ma_n = int(cfg.get('ma_months', 3))
    ma_pct = float(cfg.get('ma_trend_pct', 10))
    ma_end = prev if preview and not preview.get('full_month') else cur

    def ma_of(ks, end):
        return sum(Pf(k, month_add(end[0], end[1], -i)) for k in ks for i in range(ma_n)) / ma_n

    def win_ok(end):
        # 區間內每個月的資料都完整，才能拿來比較（例：資料從 2024/01 開始、2024/07~08 不完整）
        return all(month_complete(month_add(end[0], end[1], -i)) for i in range(ma_n))

    def ma_trend(ks):
        """回傳 (近3月月均, 前3月月均, 去年同期月均, 較前3月%, 較去年同期%, 判讀, 燈色)。"""
        before_end = month_add(ma_end[0], ma_end[1], -ma_n)
        ly_end = (ma_end[0] - 1, ma_end[1])
        now = ma_of(ks, ma_end)
        before = ma_of(ks, before_end)
        ly = ma_of(ks, ly_end)
        seq_v = pct(now, before) if win_ok(before_end) else None
        if not win_ok(ly_end):
            return now, before, ly, seq_v, None, '去年同期資料不完整', 'grey'
        yoy_v = pct(now, ly)
        if yoy_v is None:
            verdict, cls = '去年同期無生產', 'grey'
        elif yoy_v <= -ma_pct:
            verdict, cls = '趨勢下滑', 'bad'
        elif yoy_v >= ma_pct:
            verdict, cls = '趨勢成長', 'ok'
        else:
            verdict, cls = '趨勢持平', 'mid'
        return now, before, ly, seq_v, yoy_v, verdict, cls

    def ma_row(ks, label, code, core):
        now, before, ly, seq_v, yoy_v, verdict, cls = ma_trend(ks)
        months = [month_add(ma_end[0], ma_end[1], -i) for i in range(11, -1, -1)]
        return {
            '品項': label, '料號': code, 'core': core,
            'ma': round(now), 'ma_before': round(before), 'ma_ly': round(ly),
            'seq_pct': r1(seq_v), 'yoy_pct': r1(yoy_v), 'verdict': verdict, 'cls': cls,
            'labels': [mlabel(*ym) for ym in months],
            'monthly': [round(sum(Pf(k, ym) for k in ks)) for ym in months],
            'series': [round(ma_of(ks, ym)) for ym in months],
        }

    ma_rows = [ma_row(core_biz, '本業合計（排除季節性代工）' if has_oem else '本業合計', '', False)]
    ma_rows += [ma_row([k], meta[k]['品項'], meta[k]['料號'], True) for k in core_ranked]

    def win_label(end):
        st = month_add(end[0], end[1], -(ma_n - 1))
        return '%s–%s' % (mlabel(*st), mlabel(*end))
    # 配套：每筆接單下滑警示附上該品項的 3 個月趨勢，區分「單月波動」與「趨勢性衰退」
    for a in alerts:
        _n, _b, _l, _s, yv, vd, vc = ma_trend([a['key']])
        a['ma_yoy'], a['ma_verdict'], a['ma_cls'] = r1(yv), vd, vc

    ma_block = {
        'rows': ma_rows, 'n': ma_n, 'end': mlabel(*ma_end),
        'window': win_label(ma_end),
        'window_before': win_label(month_add(ma_end[0], ma_end[1], -ma_n)),
        'window_ly': win_label((ma_end[0] - 1, ma_end[1])),
    }

    # ── 綜合燈號 ──
    reasons = []
    if judge_yoy is not None and judge_yoy <= -dec:
        reasons.append('總鍋數較去年同%s下降 %.1f%%%s' % (
            '期' if preview else '月', judge_yoy, '（已排除季節性代工）' if has_oem else ''))
    red_gap = int(cfg.get('red_core_dormant_months', 1))
    core_dormant = [d for d in dormant if d['core'] and d['gap_months'] >= red_gap] if not preview else []
    if core_dormant:
        reasons.append('%d 項主力品項斷單%s' % (
            len(core_dormant), '' if red_gap <= 1 else '（連續 ≥%d 月）' % red_gap))
    red_ma = float(cfg.get('red_ma_trend_pct', 10))
    ma_total_yoy = ma_rows[0]['yoy_pct']
    if ma_total_yoy is not None and ma_total_yoy <= -red_ma:
        reasons.append('本業近 %d 個月平均較去年同期 %.1f%%（%s，趨勢性衰退）' % (
            ma_n, ma_total_yoy, ma_block['window']))
    red_pct = float(cfg.get('red_core_loss_pct', 15))
    if a_alerts and core_loss_pct is not None and core_loss_pct >= red_pct:
        reasons.append('%d 項主力品項衰退，合計流失 %s 鍋（占本業 %.1f%%，達紅燈門檻 %g%%）' % (
            len(a_alerts), '{:,}'.format(core_loss), core_loss_pct, red_pct))
    if reasons:
        light, light_label = 'red', '需注意'
    elif a_alerts:
        light, light_label = 'yellow', '觀察'
        reasons.append('%d 項主力品項衰退，合計流失 %s 鍋（占本業 %.1f%%，未達紅燈門檻 %g%%）' % (
            len(a_alerts), '{:,}'.format(core_loss), core_loss_pct or 0, red_pct))
    elif (judge_mom or 0) >= 0 and (judge_yoy or 0) >= 0:
        light, light_label = 'green', '表現良好'
        reasons.append('總量環比與同比皆未衰退，且無主力品項警示')
    else:
        light, light_label = 'yellow', '持平'
        reasons.append('總量小幅變動，無主力品項警示')
    if preview:
        light_label += '（暫定）'

    # ── 上期追蹤：上月點名的 A/B 級警示、斷單、資料異常，本月的狀態 ──
    tracking = None
    if with_tracking and prev in month_counts:
        Rp = compute(full, full, meta, month_counts, prev[0], prev[1], cfg, None, False)
        items = []

        def status_prev(k, ref_v):
            """期中預覽：只用實際值——本月至今 vs 上月同期（不做推估）。"""
            now, same = P(k, cur), P(k, prev)
            if ref_v and now >= 0.9 * ref_v:
                return '已恢復（本月至今已達基準）', 'ok'
            if now == 0:
                return '本月尚未生產', 'mid'
            if same == 0:
                return '回升中（上月同期為 0）', 'mid'
            r = (now / same - 1) * 100
            if r >= 10:
                return '回升中（較上月同期 +%.0f%%）' % r, 'mid'
            if r <= -10:
                return '持續下滑（較上月同期 %.0f%%）' % r, 'bad'
            return '與上月同期持平', 'mid'

        for a in Rp['alerts']:
            if a['grade'] not in ('A', 'B'):
                continue
            k = a['key']
            ref_v = a['pots_prev'] if a['basis'] == 'MoM' else a['pots_yoy']
            last, now = a['pots_cur'], P(k, cur)
            if preview:
                st, cls = status_prev(k, ref_v)
            elif now >= 0.9 * ref_v:
                st, cls = '已恢復', 'ok'
            elif now == 0:
                st, cls = '轉為斷單', 'bad'
            elif now > last * 1.1:
                st, cls = '回升中', 'mid'
            elif now < last * 0.9:
                st, cls = '持續下滑', 'bad'
            else:
                st, cls = '未改善', 'mid'
            items.append({'kind': '%s級衰退' % a['grade'], '料號': a['料號'], '品項': a['品項'],
                          'core': a['core'], 'tag': a.get('tag'),
                          'then': '%s 鍋（比較基準 %s，流失 %s）' % (
                              '{:,}'.format(last), '{:,}'.format(ref_v), '{:,}'.format(a['loss'])),
                          'same': round(P(k, prev)), 'now': round(now), 'status': st, 'cls': cls})
        for d in Rp['dormant']:
            k = d['key']
            if P(k, cur) > 0:
                st, cls = '恢復生產', 'ok'
            elif preview:
                st, cls = '本月尚未生產', 'mid'
            else:
                g = d['gap_months'] + 1
                st, cls = '持續斷單（連續 %d 月）' % g, 'bad'
            items.append({'kind': '斷單', '料號': d['料號'], '品項': d['品項'],
                          'core': d['core'], 'tag': d.get('tag'),
                          'then': '0 鍋（前期月均 %s）' % '{:,}'.format(d['avg_pots']),
                          'same': round(P(k, prev)), 'now': round(P(k, cur)), 'status': st, 'cls': cls})
        for x in Rp['anomalies']:
            k = x['key']
            if kg_pending:
                st, cls = '重量尚未填入，待確認', 'mid'
            elif k in anomaly_keys:
                st, cls = '仍異常', 'bad'
            elif P(k, cur) < an_min:
                st, cls = '本月未生產（無法確認）', 'mid'
            else:
                st, cls = '已正常', 'ok'
            items.append({'kind': '資料異常', '料號': x['料號'], '品項': x['品項'],
                          'core': False, 'tag': x.get('tag'),
                          'then': '%s（%s 鍋／%s kg）' % (x['kind'], '{:,}'.format(x['pots']), '{:,}'.format(x['kg'])),
                          'same': round(P(k, prev)), 'now': round(P(k, cur)), 'status': st, 'cls': cls})
        tracking = {'month': Rp['month'], 'items': items}

    return {
        'month': mlabel(Y, M), 'year': Y, 'mon': M,
        'prev_label': mlabel(*prev), 'yoy_label': mlabel(*yoy),
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'preview': preview,
        'incomplete': incomplete, 'is_latest_month': is_latest,
        'record_count': cur_count, 'median_count': median,
        'totals': {
            'pots': round(tot_pots), 'pots_prev': round(tot_pots_prev), 'pots_yoy': round(tot_pots_yoy),
            'kg': round(tot_kg), 'kg_prev': round(tot_kg_prev), 'kg_yoy': round(tot_kg_yoy),
            'items': len(active_items),
            'mom_pct': r1(mom_pct_total), 'yoy_pct': r1(yoy_pct_total),
            'kg_mom_pct': r1(pct(tot_kg, tot_kg_prev)), 'kg_yoy_pct': r1(pct(tot_kg, tot_kg_yoy)),
            'ytd': round(ytd), 'ytd_prev': round(ytd_prev), 'ytd_pct': r1(pct(ytd, ytd_prev)),
            'ytd_kg': round(ytd_kg), 'ytd_kg_prev': round(ytd_kg_prev),
        },
        'ex_oem': None if not has_oem else {
            'pots': round(ex_pots), 'prev': round(ex_prev), 'yoy': round(ex_yoy_v),
            'mom_pct': r1(ex_mom), 'yoy_pct': r1(ex_yoy_pct),
        },
        'ytd': ytd_block,
        'annotated': [
            {'料號': meta[k]['料號'], '品項': meta[k]['品項'],
             'status': o.get('status'), 'label': o.get('label', ''), 'note': o.get('note', '')}
            for k, o in sorted(ov.items(), key=lambda kv: meta[kv[0]]['料號'])
        ],
        'alerts': alerts, 'grade_counts': grade_counts, 'b_threshold': b_threshold,
        'core_loss': core_loss, 'core_loss_pct': r1(core_loss_pct),
        # 相容舊版索引頁：警示數＝A+B 級
        'real_alert_count': grade_counts['A'] + grade_counts['B'],
        'core_real': a_alerts,
        'dormant': dormant, 'growth': growth[:cfg['top_n']],
        'core_items': [{'料號': meta[k]['料號'], '品項': meta[k]['品項'], 'base12': round(baseline[k]),
                        'share': round(baseline[k] / pool_total * 100, 1) if pool_total else 0,
                        'pots_cur': round(P(k, cur)), 'pots_prev': round(P(k, prev))}
                       for k in core_ranked],
        'core_window': '%s–%s' % (mlabel(*base_window[-1]), mlabel(*base_window[0])),
        'newbies': newbies, 'returning': returning,
        'top_items': top_list, 'top_share': top_share, 'all_items': all_items,
        'categories': cats, 'same_month': same_month,
        'typical_mom': typical_mom, 'seasonal_note': seasonal_note, 'seasonal_sample': seasonal_sample,
        'trend': trend, 'waterfall': waterfall, 'anomalies': anomalies, 'tracking': tracking,
        'ma': ma_block,
        'kg_coverage': round(kg_cov, 1), 'kg_pending': kg_pending, 'nokg_pots': round(nokg_pots),
        'factory_kg_per_pot': None if factory_kpp is None else round(factory_kpp, 1),
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
    'real': '<span class="tag t-grey">重量同步下降</span>',
    'partial': '<span class="tag t-grey">重量小幅下降</span>',
    'batch': '<span class="tag t-grey">重量持平或增加</span>',
    'unknown': '<span class="tag t-grey">無重量資料</span>',
    'pending': '<span class="tag t-grey">重量尚未填入</span>',
}

GRADE_BADGE = {
    'A': '<span class="tag t-red">A 主力</span>',
    'B': '<span class="tag t-amber">B</span>',
    'C': '<span class="tag t-grey">C</span>',
    'N': '<span class="tag t-note">註</span>',
}


def item_cell(x):
    """品項欄：主力／註記標籤 + 品名 + 料號。"""
    tags = ''
    if x.get('core'):
        tags += '<span class="tag t-core">主力</span> '
    if x.get('tag'):
        tags += '<span class="tag t-note">%s</span> ' % esc(x['tag'])
    return '%s%s<div class="sub">%s</div>' % (tags, esc(x['品項']), esc(x['料號']))


def alert_rows(R, grades, empty_msg):
    rows = []
    for a in R['alerts']:
        if a['grade'] not in grades:
            continue
        if a.get('ma_verdict'):
            trend = '<span class="tag %s">%s</span>%s' % (
                MA_BADGE.get(a.get('ma_cls'), 't-grey'), esc(a['ma_verdict']),
                '<div class="sub">%s</div>' % pct_txt(a['ma_yoy']) if a.get('ma_yoy') is not None else '')
        else:
            trend = '—'
        rows.append(
            '<tr class="g-%s"><td>%s</td><td>%s</td>'
            '<td class="num">%s</td><td class="num">%s</td><td class="num">%s</td>'
            '<td class="num">%s</td><td class="num">%s</td><td class="num loss">-%s</td><td>%s</td><td>%s</td></tr>' % (
                a['grade'], GRADE_BADGE[a['grade']], item_cell(a),
                fmt(a['pots_cur']), fmt(a['pots_prev']), fmt(a['pots_yoy']),
                pct_span(a['mom_pct']), pct_span(a['yoy_pct']),
                fmt(a['loss']), trend, VERDICT_BADGE.get(a['verdict'], '')))
    return '\n'.join(rows) if rows else '<tr><td colspan="10" class="empty">%s</td></tr>' % empty_msg


def dormant_rows(R):
    if not R['dormant']:
        return '<tr><td colspan="6" class="empty">%s</td></tr>' % (
            '目前沒有「常態生產但本月尚未生產」的品項' if R.get('preview') else '本月沒有斷單品項 👍')
    rows = []
    for d in R['dormant']:
        if R.get('preview'):
            st = '<span class="tag t-grey">尚未生產</span>'
        elif d.get('suspect_lost'):
            st = '<span class="tag t-red">疑似流失</span>'
        else:
            st = '<span class="tag t-amber">斷單</span>'
        last = '%s（%s 鍋）' % (d['last_seen'], fmt(d['last_seen_pots'])) if d.get('last_seen') else '—'
        rows.append('<tr><td>%s</td><td class="num">%s</td><td class="num">%s</td>'
                    '<td class="num">%s</td><td class="num">%s 月</td><td>%s</td></tr>' % (
                        item_cell(d), d['active_months'], fmt(d['avg_pots']), last,
                        d['gap_months'], st))
    return '\n'.join(rows)


def anomaly_rows(R):
    if R.get('kg_pending'):
        return ('<tr><td colspan="7" class="empty">本月重量資料尚未填入（有重量的鍋數僅占 %s%%），'
                '暫不檢查；重量補齊後重新產生即可。</td></tr>' % R['kg_coverage'])
    if not R['anomalies']:
        return '<tr><td colspan="7" class="empty">本月沒有鍋數與重量不匹配的品項 👍</td></tr>'
    return '\n'.join(
        '<tr><td>%s</td><td class="num">%s</td><td class="num">%s</td><td class="num">%s</td>'
        '<td class="num">%s</td><td class="num">%s</td><td><span class="tag t-red">%s</span></td></tr>' % (
            item_cell(x), fmt(x['pots']), fmt(x['kg']), x['kg_per_pot'],
            '—' if x['hist_median'] is None else '%s<div class="sub">%s</div>' % (x['hist_median'], esc(x.get('basis') or '')),
            '—' if x['dev_pct'] is None else pct_txt(x['dev_pct']), esc(x['kind']))
        for x in R['anomalies'])


TRACK_BADGE = {'ok': 't-green', 'mid': 't-amber', 'bad': 't-red'}


def tracking_rows(R):
    tr = R.get('tracking')
    pv = R.get('preview')
    if not tr or not tr['items']:
        return '<tr><td colspan="%d" class="empty">上期沒有需追蹤的事項</td></tr>' % (6 if pv else 5)
    order = {'bad': 0, 'mid': 1, 'ok': 2}
    items = sorted(tr['items'], key=lambda i: order.get(i['cls'], 3))
    return '\n'.join(
        '<tr><td>%s</td><td>%s</td><td>%s</td>%s<td class="num">%s</td>'
        '<td><span class="tag %s">%s</span></td></tr>' % (
            esc(i['kind']), item_cell(i), esc(i['then']),
            '<td class="num">%s</td>' % fmt(i.get('same', 0)) if pv else '',
            fmt(i['now']), TRACK_BADGE.get(i['cls'], 't-grey'), esc(i['status']))
        for i in items)


MA_BADGE = {'ok': 't-green', 'mid': 't-grey', 'bad': 't-red', 'grey': 't-grey'}


def ma_rows_html(R):
    mb = R.get('ma')
    if not mb:
        return ''
    out = []
    for i, r in enumerate(mb['rows']):
        name = ('<strong>%s</strong>' % esc(r['品項'])) if not r['料號'] else \
            '%s<div class="sub">%s</div>' % (esc(r['品項']), esc(r['料號']))
        out.append(
            '<tr%s><td>%s</td><td class="spark"><canvas id="ma%d" width="170" height="44"></canvas></td>'
            '<td class="num"><strong>%s</strong></td><td class="num">%s</td><td class="num">%s</td>'
            '<td class="num">%s</td><td class="num">%s</td><td><span class="tag %s">%s</span></td></tr>' % (
                ' class="total"' if not r['料號'] else '', name, i,
                fmt(r['ma']), fmt(r['ma_before']), pct_span(r['seq_pct']),
                fmt(r['ma_ly']), pct_span(r['yoy_pct']),
                MA_BADGE.get(r['cls'], 't-grey'), esc(r['verdict'])))
    return '\n'.join(out)


def core_rows(R):
    return '\n'.join(
        '<tr><td class="num">%d</td><td>%s<div class="sub">%s</div></td><td class="num">%s</td>'
        '<td class="num">%s%%</td><td class="num">%s</td></tr>' % (
            i + 1, esc(c['品項']), esc(c['料號']), fmt(c['base12']), c['share'], fmt(c['pots_cur']))
        for i, c in enumerate(R.get('core_items') or []))


def growth_rows(R):
    if not R['growth']:
        return '<tr><td colspan="6" class="empty">本月沒有明顯成長的品項</td></tr>'
    return '\n'.join(
        '<tr><td>%s</td><td class="num">%s</td><td class="num">%s</td>'
        '<td class="num">%s</td><td class="num gain">+%s</td><td class="num">%s</td></tr>' % (
            item_cell(g), fmt(g['pots_cur']), fmt(g['pots_prev']),
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


def all_rows(R):
    num = lambda v: fmt(v) if v == int(v) else ('%.1f' % v)
    return '\n'.join(
        '<tr><td>%s%s<div class="sub">%s</div></td><td>%s</td><td class="num">%s</td><td class="num">%s</td>'
        '<td class="num">%s</td><td class="num">%s</td><td class="num">%s</td></tr>' % (
            ('<span class="tag t-note">%s</span> ' % esc(t['tag'])) if t.get('tag') else '',
            esc(t['品項']), esc(t['料號']), esc(t['類別']), num(t['pots']), num(t['prev']), num(t['yoy']),
            pct_span(t['mom_pct']), fmt(t['kg']) if t['kg'] else '—')
        for t in R.get('all_items') or [])


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
<title>生產鍋數月報 __MONTH____TITLE_SUFFIX__</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{--bg:#f6f7f9;--surface:#fff;--border:#e2e8f0;--text:#0f172a;--t2:#334155;--t3:#64748b;--t4:#94a3b8;
--blue:#1d4ed8;--amber:#b45309;--teal:#0f766e;--red:#dc2626;--green:#16a34a;--r:12px;
--sh:0 1px 3px rgba(15,23,42,.06);--mono:'DM Mono',ui-monospace,SFMono-Regular,Menlo,monospace;
--font:'Noto Sans TC',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font-family:var(--font);line-height:1.65;font-size:15px;}
.wrap{max-width:1240px;margin:0 auto;padding:26px 22px 60px;}
header.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;margin-bottom:20px;}
h1{font-size:26px;margin:0 0 4px;letter-spacing:-.01em;}
.sub{font-size:13.5px;color:var(--t3);}
.light{display:inline-flex;align-items:center;gap:8px;padding:10px 18px;border-radius:999px;font-weight:700;font-size:16px;}
.light.red{background:#fee2e2;color:#991b1b;} .light.yellow{background:#fef3c7;color:#92400e;} .light.green{background:#dcfce7;color:#166534;}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;box-shadow:var(--sh);margin-bottom:18px;}
.card h2{font-size:19px;margin:0 0 4px;} .card .csub{font-size:13.5px;color:var(--t3);margin-bottom:14px;line-height:1.6;}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:15px 18px;box-shadow:var(--sh);position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:linear-gradient(180deg,#1d4ed8,#6366f1);}
.kpi .lab{font-size:14px;color:var(--t3);margin-bottom:6px;}
.kpi .val{font-size:30px;font-weight:700;font-family:var(--mono);letter-spacing:-.02em;line-height:1.1;}
.kpi .dlt{font-size:13.5px;margin-top:5px;font-family:var(--mono);color:var(--t3);}
.up{color:var(--green);} .dn{color:var(--red);} .flat{color:var(--t4);}
.note{padding:12px 16px;border-radius:9px;font-size:14.5px;margin-bottom:18px;}
.note.warn{background:#fef3c7;border:1px solid #f6d68a;color:#92400e;}
.note.info{background:#eff6ff;border:1px solid #bfdbfe;color:#1e40af;}
.note.preview{background:#fff7ed;border:2px solid #fb923c;color:#9a3412;font-size:15px;line-height:1.7;}
tr.g-A{background:#fff7f7;}
.glossary summary{cursor:pointer;list-style:none;display:flex;align-items:baseline;gap:10px;}
.glossary summary::-webkit-details-marker{display:none;}
.glossary summary h2{display:inline;margin:0;}
.gl-grid{display:grid;grid-template-columns:1.1fr 1fr;gap:26px;margin-top:14px;}
.gl-grid h3{font-size:16px;margin:16px 0 6px;} .gl-grid h3:first-child{margin-top:0;}
.gl-grid p,.gl-grid li{font-size:15px;color:var(--t2);margin:4px 0;}
.gl-grid ul{margin:4px 0 8px;padding-left:22px;}
table.gl td{border-bottom:1px solid var(--border);padding:8px 8px;vertical-align:top;}
table.gl td:first-child{width:70px;}
table.ma td{vertical-align:middle;} td.spark{padding:4px 8px;width:186px;}
tr.total{background:#f8fafc;}
details.more{margin-top:12px;} details.more summary{cursor:pointer;font-size:14.5px;color:var(--blue);padding:6px 0;}
.hero{display:flex;align-items:baseline;gap:14px;flex-wrap:wrap;margin-bottom:6px;}
.hero .big{font-size:30px;font-weight:700;font-family:var(--mono);letter-spacing:-.02em;}
.hero .cmp{font-size:15px;color:var(--t3);font-family:var(--mono);}
table{width:100%;border-collapse:collapse;font-size:15px;}
th,td{padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top;}
th{background:#f8fafc;font-size:13.5px;color:var(--t3);font-weight:600;white-space:nowrap;}
td.num{text-align:right;font-family:var(--mono);white-space:nowrap;}
th.num{text-align:right;}
td .sub{font-size:12px;color:var(--t4);font-family:var(--mono);}
td.loss{color:var(--red);font-weight:600;} td.gain{color:var(--green);font-weight:600;}
tr.row-real{background:#fff7f7;}
td.empty{text-align:center;color:var(--t4);padding:20px;}
.tag{display:inline-block;font-size:12.5px;padding:2px 8px;border-radius:20px;white-space:nowrap;}
.t-red{background:#fee2e2;color:#991b1b;} .t-amber{background:#fef3c7;color:#92400e;}
.t-grey{background:#f1f5f9;color:#64748b;} .t-core{background:#e0e7ff;color:#3730a3;font-weight:700;}
.t-blue{background:#dbeafe;color:#1e40af;} .t-green{background:#dcfce7;color:#166534;}
.t-note{background:#ede9fe;color:#5b21b6;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.chart{position:relative;height:330px;}
.chart.tall{height:400px;}
.tbl-wrap{overflow-x:auto;}
.muted{color:var(--t4);font-size:14px;}
.narrative p{margin:0 0 13px;font-size:16px;line-height:1.8;color:var(--t2);}
footer{margin-top:26px;font-size:13px;color:var(--t4);line-height:1.8;}
@media(max-width:860px){.kpis{grid-template-columns:1fr 1fr;}.g2{grid-template-columns:1fr;}.gl-grid{grid-template-columns:1fr;}.wrap{padding:16px 13px 44px;}}
/* ── PDF 簡報：A4 橫式，每頁頁尾有報告名稱與頁碼；分頁由下方 script 依頁高自動安排 ── */
.top-right{display:flex;align-items:center;gap:12px;flex-wrap:wrap;}
.pdf-btn{display:inline-flex;align-items:center;gap:6px;padding:10px 16px;border-radius:999px;border:1px solid #bfdbfe;
  background:#eff6ff;color:#1d4ed8;font:inherit;font-size:15px;font-weight:600;cursor:pointer;}
.pdf-btn:hover{background:#dbeafe;}
.pdf-tip{font-size:12.5px;color:var(--t4);}
@page{size:A4 landscape;margin:10mm 10mm 12mm;
  @bottom-left{content:"生產鍋數月報 · __MONTH____TITLE_SUFFIX__";font-size:9pt;color:#94a3b8;}
  @bottom-right{content:"第 " counter(page) " / " counter(pages) " 頁";font-size:9pt;color:#94a3b8;}}
body.pdf .wrap{width:1040px;max-width:none;margin:0;padding:0;}
body.pdf .kpis{grid-template-columns:repeat(4,minmax(0,1fr));} body.pdf .g2{grid-template-columns:minmax(0,1fr) minmax(0,1fr);}
body.pdf .card{min-width:0;} body.pdf .chart canvas{max-width:100%;}
body.pdf .card{padding:14px 18px;margin-bottom:12px;} body.pdf .kpis{margin-bottom:12px;}
body.pdf table{font-size:13.5px;} body.pdf th,body.pdf td{padding:5px 8px;} body.pdf th{font-size:12.5px;}
body.pdf td .sub{font-size:11px;line-height:1.3;} body.pdf .tag{font-size:11.5px;padding:1px 7px;}
body.pdf .narrative p{font-size:15px;line-height:1.75;margin-bottom:10px;} body.pdf .card .csub{margin-bottom:10px;}
body.pdf td.spark{padding:2px 8px;}
body.pdf footer{font-size:10.5px;line-height:1.6;margin-top:6px;}
body.pdf .gl-grid{grid-template-columns:1.1fr 1fr;}
@media print{
  html,body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .no-print{display:none!important;}
  .wrap{max-width:none;padding:0;}
  .card,.kpi{box-shadow:none;}
  .card,.g2,.kpis,header.top,.note{break-inside:avoid;}
  .card.split{break-inside:auto;}
  .card h2,.card .csub,.hero{break-after:avoid;}
  tr{break-inside:avoid;} thead{display:table-header-group;}
  .tbl-wrap{overflow:visible;}
  .glossary summary .muted{display:none;}
}
</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <div>
    <h1>生產鍋數月報 · __MONTH____TITLE_SUFFIX__</h1>
    <div class="sub">__HEADER_SUB__　｜　產生於 __GEN__</div>
  </div>
  <div class="top-right">
    <div class="light __LIGHT__">__LIGHT_ICON__ __LIGHT_LABEL__</div>
    <div class="no-print"><button type="button" class="pdf-btn" id="pdfBtn" title="開啟列印視窗，目的地選「另存為 PDF」">📄 產生 PDF 簡報</button>
      <div class="pdf-tip">列印視窗選「另存為 PDF」</div></div>
  </div>
</header>

__WARN__

<div class="kpis">
  <div class="kpi"><div class="lab">__CUR_WORD__總鍋數</div><div class="val">__POTS__</div>
    <div class="dlt">環比 __POTS_MOM__ ｜ 同比 __POTS_YOY__</div>__PROJ__</div>
  <div class="kpi"><div class="lab">__CUR_WORD__總重量 (kg)</div><div class="val">__KG__</div>
    <div class="dlt">__KG_DLT__</div></div>
  <div class="kpi"><div class="lab">生產品項數</div><div class="val">__ITEMS__</div>
    <div class="dlt">前 __TOPN__ 大占 __TOPSHARE__%</div></div>
  <div class="kpi"><div class="lab">接單下滑警示（A+B 級）</div><div class="val">__ALERTS__</div>
    <div class="dlt">A __GA__ ｜ B __GB__ ｜ __DORMANT_WORD__ __DORMANT__</div></div>
</div>

<div class="card">
  <h2>本月總評</h2>
  <div class="csub">綜合判讀：__LIGHT_LABEL__ — __REASONS__</div>
  <div class="narrative">__NARRATIVE__</div>
  __EXOEM__
  __SEASON__
</div>

<div class="card glossary">
  <details>
  <summary><h2>📖 名詞說明：主力品項、A／B／C 級、衰退與斷單怎麼判定？</h2><span class="muted">（點開查看）</span></summary>
  <div class="gl-grid">
    <div>
      <h3>接單下滑警示的觸發條件</h3>
      <p>品項鍋數符合以下任一條件，就列入警示：</p>
      <ul>
        <li>較<strong>__PREV_WORD__</strong>減少 ≥ __DECLINE__%，且__PREV_WORD__ ≥ __SCALE__ 鍋</li>
        <li>較<strong>__YOY_WORD__</strong>減少 ≥ __DECLINE__%，且__YOY_WORD__ ≥ __SCALE__ 鍋</li>
      </ul>
      <p>「≥ __SCALE__ 鍋」是為了排除零星小量品項的雜訊。<strong>流失鍋數</strong>取兩者中減少較多的那一個。</p>
      <h3>警示分級（只看鍋數）</h3>
      <table class="gl">
        <tr><td><span class="tag t-red">A 主力</span></td><td>觸發警示的品項屬於<strong>主力品項</strong>（右表）。不設門檻，一律列出。</td></tr>
        <tr><td><span class="tag t-amber">B</span></td><td>非主力品項，流失 ≥ <strong>__BTHR__ 鍋</strong>（本業總鍋數的 __BPCT__%，會隨淡旺季自動調整）。</td></tr>
        <tr><td><span class="tag t-grey">C</span></td><td>非主力、流失 &lt; __BTHR__ 鍋。影響小，預設收合。</td></tr>
        <tr><td><span class="tag t-note">註</span></td><td>特殊品項（見頁尾「品項註記」）：季節性代工另列、不分級、不影響燈號；標「不納入分析」者所有數據皆已排除。</td></tr>
      </table>
      <h3>__DORMANT_TITLE__</h3>
      <p>__DORMANT_SUB__</p>
      <h3>綜合燈號</h3>
      <p>🔴 符合任一：本業較__YOY_WORD__減少 ≥ __DECLINE__%；本業<strong>近 3 個月平均</strong>較去年同期減少 ≥ __REDMA__%（趨勢性衰退）；主力品項斷單__REDGAP__；A 級流失合計 ≥ 本業總鍋數的 __REDPCT__%<br>
         🟡 有 A 級衰退但未達紅燈，或總量小幅衰退　🟢 總量未衰退且無主力警示</p>
      <p class="muted">「本業」＝全廠扣除季節性代工品項。</p>
    </div>
    <div>
      <h3>主力品項（__CORE_N__ 項）</h3>
      <p>依<strong>前 12 個月（__CORE_WINDOW__）累計鍋數</strong>排名前 __CORE_N__ 名，排除季節性代工、已下市與不納入分析的品項。
         用前 12 個月而不用本月排名，才不會因為本月衰退就跌出主力名單。每月自動更新。</p>
      <div class="tbl-wrap"><table>
        <thead><tr><th class="num">#</th><th>品項</th><th class="num">12 個月鍋數</th><th class="num">占比</th><th class="num">__CUR_COL__</th></tr></thead>
        <tbody>__CORE_ROWS__</tbody>
      </table></div>
    </div>
  </div>
  </details>
</div>

<div class="card">
  <h2>上期追蹤 — __TRACK_MONTH__ 點名事項</h2>
  <div class="csub">__TRACK_SUB__</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>類型</th><th>品項</th><th>上期狀況</th>__TRACK_SAME_TH__<th class="num">__TRACK_NOW__</th><th>本月狀態</th></tr></thead>
    <tbody>__TRACK_ROWS__</tbody>
  </table></div>
</div>

<div class="card">
  <h2>環比變動拆解</h2>
  <div class="csub">__WF_SUB__</div>
  <div class="chart tall"><canvas id="wfChart"></canvas></div>
</div>

<div class="card">
  <h2>整體業務量趨勢</h2>
  <div class="csub">近 __TRENDN__ 個月：柱狀為鍋數、折線為半成品重量(kg)。以鍋數為主要指標，重量為輔助參考。__TREND_NOTE__</div>
  <div class="chart tall"><canvas id="trendChart"></canvas></div>
</div>

<div class="card">
  <h2>年度累計（YTD）</h2>
  <div class="hero"><div class="big">__YTD_POTS__ 鍋</div><div class="cmp">較去年同期 __YTD_PCT__（去年 __YTD_PREV__ 鍋）</div></div>
  <div class="csub">__YTD_NOTE__</div>
  <div class="chart"><canvas id="ytdChart"></canvas></div>
</div>

<div class="g2">
  <div class="card"><h2>__CMP_TITLE__</h2><div class="csub">鍋數與重量雙軌對照</div>
    <div class="chart"><canvas id="cmpChart"></canvas></div></div>
  <div class="card"><h2>季節性對照</h2><div class="csub">歷年同月份（__MONTHNUM__月__SEASON_PERIOD__）總鍋數，用以區分季節性淡季與實質衰退。__SEASON_SAMPLE__</div>
    <div class="chart"><canvas id="seasonChart"></canvas></div></div>
</div>

<div class="card">
  <h2>接單下滑警示 — A／B 級流失鍋數</h2>
  <div class="csub">依流失鍋數由大到小，最多列 __TOPN__ 項（C 級與季節性代工不列入圖表）</div>
  <div class="chart tall"><canvas id="declineChart"></canvas></div>
</div>

<div class="card">
  <h2>接單下滑警示明細（分級）</h2>
  <div class="csub"><strong>A 級</strong>＝主力品項｜<strong>B 級</strong>＝非主力且流失 ≥ __BTHR__ 鍋｜<strong>C 級</strong>＝其餘，預設收合｜<strong>註</strong>＝季節性代工。判定方式見上方「📖 名詞說明」。<strong>3 個月趨勢</strong>：該品項近 3 個完整月份平均較去年同期（趨勢下滑＝趨勢性衰退，持平／成長＝多為單月排程波動）。「重量參考」欄僅供參考，不影響分級。</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>等級</th><th>品項</th><th class="num">__CUR_COL__</th><th class="num">__PREV_COL__</th><th class="num">__YOY_COL__</th>
      <th class="num">環比</th><th class="num">同比</th><th class="num">流失鍋數</th><th>3 個月趨勢</th><th>重量參考</th></tr></thead>
    <tbody>__ALERT_ROWS__</tbody>
  </table></div>
  __C_DETAILS__
</div>

<div class="card">
  <h2>主力品項 3 個月趨勢（移動平均）</h2>
  <div class="csub">單月產量受排程影響波動大，改看<strong>近 3 個月平均月鍋數</strong>（__MA_WIN__）判斷趨勢：
    與去年同期 3 個月（__MA_WIN_LY__）相比 ≤ −__MA_PCT__% 為「趨勢下滑」、≥ +__MA_PCT__% 為「趨勢成長」，可同時排除季節性。__MA_NOTE__
    小圖：灰色長條為各月鍋數、藍線為 3 個月移動平均（近 12 個月）。<strong>本業合計</strong>較去年同期 ≤ −__REDMA__% 時亮紅燈（趨勢性衰退）；個別品項的趨勢僅供參考，不影響分級。</div>
  <div class="tbl-wrap"><table class="ma">
    <thead><tr><th>品項</th><th>近 12 個月走勢</th><th class="num">近 3 月<br>月均</th><th class="num">前 3 月<br>月均</th>
      <th class="num">較前 3 月</th><th class="num">去年同期<br>月均</th><th class="num">較去年同期</th><th>趨勢判讀</th></tr></thead>
    <tbody>__MA_ROWS__</tbody>
  </table></div>
</div>

<div class="card">
  <h2>__DORMANT_TITLE__</h2>
  <div class="csub">__DORMANT_SUB__</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">前期活躍月數</th><th class="num">平均月鍋數</th><th class="num">最後生產月</th><th class="num">連續未生產</th><th>狀態</th></tr></thead>
    <tbody>__DORMANT_ROWS__</tbody>
  </table></div>
</div>

<div class="card">
  <h2>資料異常 — 請生管單位說明</h2>
  <div class="csub">條件：本月鍋數 ≥ __AN_MIN__ 且 (1)「有鍋數但重量為 0」；(2) 每鍋重量偏離該品項前 12 個月中位數 ±__AN_PCT__% 以上；或 (3) 歷史不足 __AN_HIST__ 個月的品項，每鍋重量低於全廠平均的 __AN_LOW__% 或高於 __AN_HIGH__%。本報告不推測原因，請生管單位說明。</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">本月鍋數</th><th class="num">本月重量 kg</th><th class="num">每鍋 kg</th>
      <th class="num">比較基準 kg</th><th class="num">偏離</th><th>異常類型</th></tr></thead>
    <tbody>__ANOMALY_ROWS__</tbody>
  </table></div>
</div>

<div class="g2">
  <div class="card"><h2>成長品項 Top __TOPN__</h2><div class="csub">依鍋數增加量排序</div>
    <div class="chart"><canvas id="growthChart"></canvas></div></div>
  <div class="card"><h2>產量集中度（Pareto）</h2><div class="csub">本月前 __TOPN__ 大品項鍋數與累積佔比</div>
    <div class="chart"><canvas id="paretoChart"></canvas></div></div>
</div>

<div class="card">
  <h2>成長品項明細</h2>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">__CUR_COL__</th><th class="num">__PREV_COL__</th><th class="num">__YOY_COL__</th>
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
  <h2>本月產量前 __TOPN__ 大品項</h2>
  <div class="csub">依本月鍋數排名（與「主力品項」不同：主力品項是依前 12 個月累計排名，見上方名詞說明）</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th class="num">鍋數</th><th class="num">佔比</th><th class="num">累積佔比</th></tr></thead>
    <tbody>__TOP_ROWS__</tbody>
  </table></div>
</div>

<div class="card glossary">
  <details>
  <summary><h2>附錄：本月全品項明細（__ALL_N__ 項）</h2><span class="muted">（點開查看；產生 PDF 時自動展開）</span></summary>
  <div class="csub" style="margin-top:8px">本月有生產的全部品項，依鍋數由多到少排列，含上方圖表與清單未列出的小量品項。重量為半成品重量(kg)，「—」表示尚未填寫。</div>
  <div class="tbl-wrap"><table>
    <thead><tr><th>品項</th><th>類別</th><th class="num">__CUR_COL__</th><th class="num">__PREV_COL__</th><th class="num">__YOY_COL__</th>
      <th class="num">環比</th><th class="num">重量 kg</th></tr></thead>
    <tbody>__ALL_ROWS__</tbody>
  </table></div>
  </details>
</div>

<footer>
  資料來源：<code>data/latest.xlsx</code> →「資料總表」　｜　分析邏輯與門檻：<code>docs/monthly-report-logic.md</code><br>
  本報告聚焦業務量（鍋數為主、半成品重量交叉驗證）；製成率／品質分析由「產品工時及製成率統計系統」負責，不在本報告範圍。<br>
  警示門檻：降幅 ≥__DECLINE__%、基準量 ≥__SCALE__ 鍋；B 級流失 ≥ 本業總鍋數 __BPCT__%　｜　斷單：前 __LOOKBACK__ 個月 ≥__MINACTIVE__ 月有生產而本月為 0，連續 ≥__LOSTM__ 月列為疑似流失<br>
  燈號：🔴 本業同比 ≤ −__DECLINE__%、本業近 3 個月平均較去年同期 ≤ −__REDMA__%、主力品項斷單__REDGAP__，或主力品項流失合計 ≥ 本業 __REDPCT__%　｜　🟡 有主力品項衰退但未達紅燈　｜　🟢 總量未衰退且無主力警示
  __ANNOTATED__
</footer>
</div>

<script>
const R = __DATA__;
const C = {blue:'#1d4ed8', amber:'#b45309', teal:'#0f766e', red:'#dc2626', green:'#16a34a', grey:'#94a3b8'};
const grid = {color:'#eef2f7'}, tick = {color:'#475569', font:{size:13}};
Chart.defaults.font.size = 13;
const baseOpts = (extra) => Object.assign({
  responsive:true, maintainAspectRatio:false,
  interaction:{mode:'index', intersect:false},
  plugins:{legend:{labels:{color:'#334155', usePointStyle:true, pointStyle:'circle', font:{size:13.5}}}}
}, extra || {});
const hOpts = () => ({
  responsive:true, maintainAspectRatio:false, indexAxis:'y',
  interaction:{mode:'index', intersect:false, axis:'y'},
  plugins:{legend:{labels:{color:'#334155', usePointStyle:true, pointStyle:'circle', font:{size:13.5}}}},
  scales:{x:{grid:grid, ticks:tick, beginAtZero:true}, y:{grid:{display:false}, ticks:{color:'#0f172a', font:{size:13}}}}
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
    y:{position:'left', grid:grid, ticks:tick, title:{display:true, text:'鍋數', color:'#64748b', font:{size:13}}},
    y1:{position:'right', grid:{display:false}, ticks:tick, title:{display:true, text:'kg', color:'#64748b', font:{size:13}}}
  }})
});

// 2. 三期比較
new Chart(document.getElementById('cmpChart'), {
  type:'bar',
  data:{labels:R.preview ? ['本月至今 '+R.preview.cur_period, '上月同期 '+R.preview.prev_period, '去年同期 '+R.preview.yoy_period]
                         : ['本月 '+R.month, '上月 '+R.prev_label, '去年同月 '+R.yoy_label], datasets:[
    {label:'鍋數', data:[R.totals.pots, R.totals.pots_prev, R.totals.pots_yoy], backgroundColor:C.blue+'cc', borderRadius:5, yAxisID:'y'},
    {label:'重量 (kg)', data:[R.totals.kg, R.totals.kg_prev, R.totals.kg_yoy], backgroundColor:C.teal+'99', borderRadius:5, yAxisID:'y1'}
  ]},
  options: baseOpts({scales:{
    x:{grid:{display:false}, ticks:tick},
    y:{position:'left', grid:grid, ticks:tick, title:{display:true, text:'鍋數', color:'#64748b', font:{size:13}}},
    y1:{position:'right', grid:{display:false}, ticks:tick, title:{display:true, text:'kg', color:'#64748b', font:{size:13}}}
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
const dec = R.alerts.filter(a=>a.grade==='A'||a.grade==='B').slice(0, R.config.top_n);
new Chart(document.getElementById('declineChart'), {
  type:'bar',
  data:{labels:dec.map(a=>shortName(a['品項'])), datasets:[
    {label:R.preview?'較上月同期流失鍋數':'較上月流失鍋數', data:dec.map(a=>Math.max(0, a.pots_prev-a.pots_cur)), backgroundColor:C.red+'bb', borderRadius:4},
    {label:R.preview?'較去年同期流失鍋數':'較去年同月流失鍋數', data:dec.map(a=>Math.max(0, a.pots_yoy-a.pots_cur)), backgroundColor:C.amber+'99', borderRadius:4}
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

// 8. 環比瀑布（浮動長條：總量為藍、減少為紅、增加為綠）
(function(){
  const W = R.waterfall, labels = [], data = [], colors = [], deltas = [];
  let run = W.start, lo = Math.min(W.start, W.end), hi = Math.max(W.start, W.end);
  W.steps.forEach(st => { run += st.delta; lo = Math.min(lo, run); hi = Math.max(hi, run); });
  run = W.start;
  // 縱軸不從 0 開始，否則小幅增減會看不見（總量長條以截斷方式呈現，說明文字已註明）
  const pad = Math.max(50, (hi - lo) * 0.15), step = Math.pow(10, Math.floor(Math.log10(Math.max(1, hi - lo + 2*pad))));
  const yMin = Math.max(0, Math.floor((lo - pad) / step) * step);
  labels.push(R.preview ? '上月同期' : '上月 '+R.prev_label); data.push([0, W.start]); colors.push(C.blue+'cc'); deltas.push(null);
  W.steps.forEach(st => {
    const next = run + st.delta;
    labels.push(shortName(st.label)); data.push([Math.min(run, next), Math.max(run, next)]);
    colors.push(st.delta < 0 ? C.red+'bb' : C.green+'bb'); deltas.push(st.delta);
    run = next;
  });
  labels.push(R.preview ? '本月至今' : '本月 '+R.month); data.push([0, W.end]); colors.push(C.blue+'cc'); deltas.push(null);
  const fullNames = [labels[0]].concat(W.steps.map(st => st.label + (st.code ? '（'+st.code+'）' : ''))).concat([labels[labels.length-1]]);
  new Chart(document.getElementById('wfChart'), {
    type:'bar',
    data:{labels:labels, datasets:[{label:'鍋數', data:data, backgroundColor:colors, borderRadius:3, borderSkipped:false}]},
    options:{responsive:true, maintainAspectRatio:false,
      plugins:{legend:{display:false}, tooltip:{callbacks:{
        title:(items)=>fullNames[items[0].dataIndex],
        label:(ctx)=>{const d=deltas[ctx.dataIndex]; const v=ctx.raw;
          return d===null ? '合計 '+(v[1]).toLocaleString()+' 鍋' : (d>0?'+':'')+d.toLocaleString()+' 鍋';}}}},
      scales:{x:{grid:{display:false}, ticks:Object.assign({maxRotation:60, autoSkip:false}, tick)},
              y:{grid:grid, ticks:tick, min:yMin, title:{display:true, text:'鍋數', color:'#64748b', font:{size:13}}}}}
  });
})();

// 9. YTD 累計（今年 vs 去年；有季節性代工時主線為排除後）
(function(){
  const Y = R.ytd, S = Y.series, ds = [
    {label:(Y.has_oem?'今年（排除季節性代工）':'今年'), data:S.this_ex, borderColor:C.blue, backgroundColor:C.blue, borderWidth:2, pointRadius:3, tension:.2},
    {label:(Y.has_oem?'去年（排除季節性代工）':'去年'), data:S.last_ex, borderColor:C.grey, backgroundColor:C.grey, borderWidth:2, pointRadius:3, tension:.2}
  ];
  if (Y.has_oem) ds.push({label:'今年（含季節性代工）', data:S.this_all, borderColor:C.amber, backgroundColor:C.amber,
    borderWidth:2, borderDash:[5,4], pointRadius:2, tension:.2});
  new Chart(document.getElementById('ytdChart'), {type:'line', data:{labels:S.labels, datasets:ds},
    options: baseOpts({scales:{x:{grid:{display:false}, ticks:tick},
      y:{grid:grid, ticks:tick, beginAtZero:true, title:{display:true, text:'累計鍋數', color:'#64748b', font:{size:13}}}}})});
})();

// 10. 主力品項 3 個月移動平均小圖（灰長條＝月鍋數，藍線＝移動平均）
(R.ma ? R.ma.rows : []).forEach((r, i) => {
  const el = document.getElementById('ma' + i);
  if (!el) return;
  new Chart(el, {
    data:{labels:r.labels, datasets:[
      {type:'line', label:'3 個月平均', data:r.series, borderColor:C.blue, borderWidth:2, pointRadius:0, tension:.3},
      {type:'bar', label:'月鍋數', data:r.monthly, backgroundColor:C.grey+'66', borderRadius:2}
    ]},
    options:{responsive:false, animation:false, plugins:{legend:{display:false},
      tooltip:{mode:'index', intersect:false, titleFont:{size:12}, bodyFont:{size:12}}},
      scales:{x:{display:false}, y:{display:false, beginAtZero:true}}}
  });
});

// 7. 類別
const cats = Object.keys(R.categories);
new Chart(document.getElementById('catChart'), {
  type:'bar',
  data:{labels:cats, datasets:[
    {label:R.preview?'本月至今':'本月', data:cats.map(c=>R.categories[c].pots), backgroundColor:C.blue+'cc', borderRadius:5},
    {label:R.preview?'上月同期':'上月', data:cats.map(c=>R.categories[c].prev), backgroundColor:C.grey+'99', borderRadius:5},
    {label:R.preview?'去年同期':'去年同月', data:cats.map(c=>R.categories[c].yoy), backgroundColor:C.amber+'99', borderRadius:5}
  ]},
  options: baseOpts({scales:{x:{grid:{display:false}, ticks:tick}, y:{grid:grid, ticks:tick, beginAtZero:true}}})
});

// 11. PDF 簡報：列印前切成 A4 橫式寬度、展開所有收合區塊、圖表以高解析度重繪，
//     再依頁高把區塊排進每一頁（放得下就同頁，放不下就換頁；圖表不會被切開，長表格可跨頁續接）。
(function(){
  const PAGE_H = 700;   // A4 橫式扣除邊界後的可用高度（px，96dpi）
  let reopened = [];
  const charts = () => Object.values(Chart.instances || {});
  function redraw(dpr){
    charts().forEach(c => {
      c.options.devicePixelRatio = dpr;
      if (c.options.responsive) c.resize(); else c.resize(c.width, c.height);
    });
  }
  function layout(){
    if (document.body.classList.contains('pdf')) return;
    document.body.classList.add('pdf');
    reopened = Array.from(document.querySelectorAll('details:not([open])'));
    reopened.forEach(d => { d.open = true; });
    redraw(3);
    const blocks = Array.from(document.querySelectorAll('.wrap > *'))
      .filter(el => el.tagName !== 'SCRIPT' && el.offsetHeight > 0);
    const pages = [[]];
    let used = 0, flowing = false;
    blocks.forEach(el => {
      el.style.breakBefore = '';
      // 不含大圖表的卡片（文字、表格）可以跨頁續接
      const split = el.classList.contains('card') && !el.querySelector('.chart');
      el.classList.toggle('split', split);
      const cs = getComputedStyle(el);
      const h = el.getBoundingClientRect().height + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
      if (el.tagName === 'FOOTER') return;   // 頁尾說明一律接在最後一頁內容後面
      if (used > 0 && used + h > PAGE_H) {
        if (split && PAGE_H - used > PAGE_H * 0.35) {
          flowing = true;            // 表格從本頁下半開始、續接到下一頁
        } else {
          el.style.breakBefore = 'page';
          pages.push([]); used = 0; flowing = false;
        }
      }
      pages[pages.length - 1].push({el: el, h: h});
      if (flowing || h > PAGE_H) {   // 跨頁的區塊：之後的頁面不做放大，避免推擠
        pages[pages.length - 1].flow = true;
        pages.push([]); pages[pages.length - 1].flow = true;
        used = (used + h) % PAGE_H; flowing = false;
      } else {
        used += h;
      }
    });
    // 同一頁的剩餘空間平均分給該頁的圖表，讓圖表放大、頁面不留大片空白
    pages.forEach(pg => {
      if (pg.flow || !pg.length) return;
      const left = PAGE_H - 12 - pg.reduce((a, b) => a + b.h, 0);
      const grow = pg.filter(b => b.el.querySelector('.chart'));
      if (left < 30 || !grow.length) return;
      const add = Math.min(260, left / grow.length);
      grow.forEach(b => b.el.querySelectorAll('.chart').forEach(c => {
        c.style.height = (c.getBoundingClientRect().height + add) + 'px';
      }));
    });
    redraw(3);
  }
  function restore(){
    if (!document.body.classList.contains('pdf')) return;
    document.body.classList.remove('pdf');
    reopened.forEach(d => { d.open = false; });
    reopened = [];
    document.querySelectorAll('.wrap > *').forEach(el => { el.style.breakBefore = ''; el.classList.remove('split'); });
    document.querySelectorAll('.chart').forEach(c => { c.style.height = ''; });
    redraw(window.devicePixelRatio || 1);
  }
  window.__pdfLayout = layout; window.__pdfRestore = restore;
  window.addEventListener('beforeprint', layout);
  window.addEventListener('afterprint', restore);
  const btn = document.getElementById('pdfBtn');
  if (btn) btn.addEventListener('click', () => { layout(); setTimeout(() => window.print(), 250); });
})();
</script>
</body>
</html>
"""

LIGHT_ICON = {'red': '🔴', 'yellow': '🟡', 'green': '🟢'}


def build_html(R, narrative):
    t = R['totals']
    cfg = R['config']
    pv = R.get('preview')
    gc = R['grade_counts']
    warn = ''
    if pv:
        warn = ('<div class="note preview">⏳ <strong>期中預覽 — %s，非正式月報</strong>：'
                '資料截至 <strong>%s</strong>（第 %d 天／共 %d 天）。所有比較皆採「同期」口徑'
                '（本月 %s 對比上月 %s、去年 %s），數字皆為實際值、不做推估。'
                '斷單與燈號皆為暫定，%s。%s</div>' % (
                    '本月資料尚未補齊' if pv.get('full_month') else '本月尚未結束',
                    pv['asof'], pv['day'], pv['days_in_month'], pv['cur_period'],
                    pv['prev_period'], pv['yoy_period'],
                    '待生管補齊資料、無空缺後產生正式月報' if pv.get('full_month') else '正式結論以下月初的正式月報為準',
                    ('<br>⚠️ 目前僅 %d 天資料，單日排程差異就會讓比較大幅波動，判讀請保守。' % pv['day']
                     if pv['day'] < 10 else '')
                    + ('<br>⚖️ 本月重量尚未填入（有重量的鍋數僅占 %s%%），重量相關數字暫不具參考性。' % R['kg_coverage']
                       if R.get('kg_pending') else
                       '<br>⚖️ 尚有 %s 鍋未填重量（%s），重量數字偏低、暫不比較；這些批次不列入資料異常。' % (
                           fmt(R['nokg_pots']), esc(pv.get('nokg_days') or ''))
                       if R.get('nokg_pots') else '')
                    + ('<br>📝 <strong>資料尚有空缺：%d 筆有鍋數，但半成品數（實際）%d 筆、成品數（實際）%d 筆未填</strong>'
                       '（%s），請生管補齊；補齊後才產生正式月報。' % (
                           pv['blank']['rows'], pv['blank']['semi'], pv['blank']['fg'], pv['blank']['days'])
                       if pv.get('blank') else '')
                    + ''.join(('<br>🏖️ %s 連續 %d 天沒有生產紀錄：<strong>計畫性停產（%s）</strong>，'
                               '同期比較會因生產天數較少而偏低。' % (g['range'], g['days'], esc(g['planned'])))
                              if g.get('planned') else
                              ('<br>📭 <strong>%s 連續 %d 天沒有任何生產紀錄</strong>（過去一年正常月份最長 4 天），'
                               '請確認是停產還是資料尚未登錄——若為漏登，本預覽的同期比較會偏低。' % (g['range'], g['days']))
                              for g in pv.get('gaps') or [])))
    elif R.get('kg_pending'):
        warn = ('<div class="note warn">⚖️ <strong>本月重量資料尚未填入</strong>：有重量的鍋數僅占 %s%%，'
                '重量相關數字與資料異常檢查暫不具參考性。</div>' % R['kg_coverage'])
    if not pv and R['incomplete']:
        warn += ('<div class="note warn">⚠️ <strong>%s 的資料可能尚未補齊</strong>：本月僅 %d 筆紀錄，'
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
                         'excluded': '資料有誤，所有月份數據皆不納入分析（總量、YTD、趨勢、警示皆已排除）',
                         'seasonal_oem': '不列主力、不參與分級與燈號、YTD 另計'}
        parts = ['<strong>%s</strong> %s（%s）%s' % (
            esc(a['label']), esc(a['品項']), esc(a['料號']),
            STATUS_EFFECT.get(a['status'], '')) for a in R['annotated']]
        ann = '<br>品項註記：' + '　｜　'.join(parts)

    season = ''
    if R['seasonal_note']:
        season = '<div class="note info" style="margin:14px 0 0">📅 <strong>季節性判讀：</strong>%s<br><span style="font-size:12px">%s</span></div>' % (
            esc(R['seasonal_note']), esc(R.get('seasonal_sample') or ''))

    # 比較口徑用詞（期中預覽改為「同期」）
    if pv:
        cur_col, prev_col, yoy_col = '本月至今', '上月同期', '去年同期'
        prev_word, yoy_word = '上月同期', '去年同期'
        header_sub = '期中預覽｜資料截至 %s　｜　與上月同期（%s）、去年同期（%s）比較' % (
            pv['asof'], pv['prev_period'], pv['yoy_period'])
        cur_word = '本月至今'
        proj = ''
        cmp_title = '本月至今 vs 上月同期 vs 去年同期'
        season_period = '' if pv['full_month'] else '，1–%d 日同期' % pv['day']
    else:
        cur_col, prev_col, yoy_col = '本月', '上月', '去年同月'
        prev_word, yoy_word = '上月', '去年同月'
        header_sub = '業務量／接單趨勢分析　｜　與 %s（上月）、%s（去年同月）比較' % (R['prev_label'], R['yoy_label'])
        cur_word = '本月'
        proj = ''
        cmp_title = '本月 vs 上月 vs 去年同月'
        season_period = ''

    # C 級收合
    c_details = ''
    if gc['C']:
        c_details = ('<details class="more"><summary>展開 C 級 %d 項（非主力、流失 &lt; %s 鍋）</summary>'
                     '<div class="tbl-wrap"><table><thead><tr><th>等級</th><th>品項</th><th class="num">%s</th>'
                     '<th class="num">%s</th><th class="num">%s</th><th class="num">環比</th><th class="num">同比</th>'
                     '<th class="num">流失鍋數</th><th>3 個月趨勢</th><th>重量參考</th></tr></thead><tbody>%s</tbody></table></div></details>' % (
                         gc['C'], fmt(R['b_threshold']), cur_col, prev_col, yoy_col,
                         alert_rows(R, ('C',), '')))

    # 上期追蹤
    tr = R.get('tracking')
    if tr:
        track_month = tr['month']
        track_sub = ('依現行規則，%s 列為 A／B 級衰退、斷單或資料異常的品項，在本月的狀態（需處理的排在最前面）。'
                     '「已恢復」＝回到當時比較基準的 90%% 以上。' % tr['month'])
        if pv:
            track_sub = ('依現行規則，%s 列為 A／B 級衰退、斷單或資料異常的品項，在本月的狀態（需處理的排在最前面）。'
                         '本月尚未結束，以「本月至今」與「上月同期」的實際鍋數比較；'
                         '本月至今已達當時比較基準 90%% 以上者列為「已恢復」。' % tr['month'])
    else:
        track_month, track_sub = R['prev_label'], '資料中沒有上月紀錄，無法追蹤。'

    # 瀑布說明
    W = R['waterfall']
    wf_sub = '%s %s 鍋 → %s %s 鍋：減少合計 %s 鍋、增加合計 +%s 鍋。列出增減最大的各 %d 項，其餘合併；縱軸未從 0 開始以便看清增減，游標移到長條上可看完整品名與數字。' % (
        '上月同期' if pv else '上月', fmt(W['start']), cur_word, fmt(W['end']),
        fmt(W['dec_total']), fmt(W['inc_total']), int(cfg.get('waterfall_items', 5)))

    # YTD
    yb = R['ytd']
    if yb['has_oem']:
        ytd_note = ('已排除季節性代工品項 %s：今年 %s 鍋、去年同期 %s 鍋。若包含，今年累計為 %s 鍋（較去年同期 %s）。'
                    '代工量體大且間歇，排除後較能反映本業的年度走勢。' % (
                        esc(yb['oem_names']), fmt(yb['oem_pots']), fmt(yb['oem_pots_prev']),
                        fmt(yb['all_pots']), pct_txt(yb['all_pct'])))
    else:
        ytd_note = '1 月至本月累計鍋數，今年 vs 去年同期。'
    if pv and not pv['full_month']:
        ytd_note += '本月以 %s 同期計算。' % pv['cur_period']

    if pv:
        dormant_title = '截至目前尚未生產（非斷單）'
        dormant_sub = ('前 %s 個月中有 ≥%s 個月正常生產，但本月截至 %s 尚未生產。本月尚未結束，'
                       '不列為斷單、不影響燈號；開會時可確認是否已排程。' % (
                           cfg['dormant_lookback'], cfg['dormant_min_active'], pv['asof']))
        dormant_word = '尚未生產'
    else:
        dormant_title = '斷單警示'
        dormant_sub = ('前 %s 個月中有 ≥%s 個月正常生產，但本月完全沒有生產。連續未生產 ≥ %s 個月列為「疑似流失」。' % (
            cfg['dormant_lookback'], cfg['dormant_min_active'], cfg.get('dormant_lost_months', 3)))
        dormant_word = '斷單'

    trend_note = ''
    if pv:
        trend_note = '（標 * 的本月為截至 %s，尚未完整）' % pv['asof']

    empty_ab = '本期沒有 A／B 級警示 👍' + ('（C 級 %d 項已收合於下方）' % gc['C'] if gc['C'] else '')
    repl = {
        '__MONTH__': R['month'], '__PREV__': R['prev_label'], '__YOY__': R['yoy_label'],
        '__TITLE_SUFFIX__': '（期中預覽）' if pv else '',
        '__HEADER_SUB__': header_sub,
        '__GEN__': R['generated_at'], '__MONTHNUM__': str(R['mon']),
        '__LIGHT__': R['light'], '__LIGHT_LABEL__': R['light_label'],
        '__LIGHT_ICON__': LIGHT_ICON.get(R['light'], ''),
        '__REASONS__': esc('；'.join(R['light_reasons'])),
        '__WARN__': warn, '__SEASON__': season,
        '__EXOEM__': exo, '__ANNOTATED__': ann,
        '__CUR_WORD__': cur_word, '__PROJ__': proj,
        '__POTS__': fmt(t['pots']), '__KG__': '—' if R.get('kg_pending') else fmt(t['kg']), '__ITEMS__': str(t['items']),
        '__POTS_MOM__': pct_span(t['mom_pct']), '__POTS_YOY__': pct_span(t['yoy_pct']),
        '__KG_DLT__': ('重量尚未填入' if R.get('kg_pending') else
                       '重量補登中（已填 %s%% 鍋數），暫不比較' % R['kg_coverage']
                       if pv and R.get('nokg_pots') else
                       '環比 %s ｜ 同比 %s' % (pct_span(t['kg_mom_pct']), pct_span(t['kg_yoy_pct']))),
        '__ALERTS__': str(gc['A'] + gc['B']), '__GA__': str(gc['A']), '__GB__': str(gc['B']),
        '__DORMANT__': str(len(R['dormant'])), '__DORMANT_WORD__': dormant_word,
        '__DORMANT_TITLE__': dormant_title, '__DORMANT_SUB__': dormant_sub,
        '__TOPSHARE__': str(R['top_share']), '__TOPN__': str(cfg['top_n']),
        '__DECLINE__': str(cfg['decline_pct']), '__SCALE__': str(cfg['min_scale_pots']),
        '__BTHR__': fmt(R['b_threshold']), '__BPCT__': '%g' % float(cfg.get('b_grade_pct', 1)),
        '__REDPCT__': '%g' % float(cfg.get('red_core_loss_pct', 15)),
        '__REDMA__': '%g' % float(cfg.get('red_ma_trend_pct', 10)),
        '__REDGAP__': '' if int(cfg.get('red_core_dormant_months', 1)) <= 1 else '（連續 ≥%d 月）' % int(cfg['red_core_dormant_months']),
        '__LOSTM__': str(cfg.get('dormant_lost_months', 3)),
        '__AN_MIN__': '%g' % float(cfg.get('anomaly_min_pots', 5)),
        '__AN_PCT__': '%g' % float(cfg.get('anomaly_dev_pct', 50)),
        '__AN_HIST__': str(cfg.get('anomaly_min_history', 3)),
        '__AN_LOW__': '%g' % float(cfg.get('anomaly_new_low_pct', 20)),
        '__AN_HIGH__': '%g' % float(cfg.get('anomaly_new_high_pct', 300)),
        '__LOOKBACK__': str(cfg['dormant_lookback']), '__MINACTIVE__': str(cfg['dormant_min_active']),
        '__TRENDN__': str(len(R['trend'])), '__TREND_NOTE__': trend_note,
        '__CUR_COL__': cur_col, '__PREV_COL__': prev_col, '__YOY_COL__': yoy_col,
        '__PREV_WORD__': prev_word, '__YOY_WORD__': yoy_word, '__CMP_TITLE__': cmp_title,
        '__SEASON_PERIOD__': season_period, '__SEASON_SAMPLE__': esc(R.get('seasonal_sample') or ''),
        '__TRACK_MONTH__': track_month, '__TRACK_SUB__': track_sub,
        '__TRACK_NOW__': '本月至今' if pv else '本月',
        '__TRACK_SAME_TH__': '<th class="num">上月同期</th>' if pv else '',
        '__CORE_ROWS__': core_rows(R), '__CORE_WINDOW__': esc(R.get('core_window') or ''),
        '__CORE_N__': str(len(R.get('core_items') or [])),
        '__MA_ROWS__': ma_rows_html(R),
        '__MA_WIN__': esc(R['ma']['window']), '__MA_WIN_LY__': esc(R['ma']['window_ly']),
        '__MA_PCT__': '%g' % float(cfg.get('ma_trend_pct', 10)),
        '__MA_NOTE__': ('本月尚未結束，只計算到上個完整月份（%s）。' % esc(R['ma']['end'])) if pv and not pv['full_month'] else '',
        '__TRACK_ROWS__': tracking_rows(R),
        '__WF_SUB__': wf_sub,
        '__YTD_POTS__': fmt(yb['pots']), '__YTD_PCT__': pct_span(yb['pct']),
        '__YTD_PREV__': fmt(yb['pots_prev']), '__YTD_NOTE__': ytd_note,
        '__NARRATIVE__': render_narrative(narrative) if narrative else render_narrative(None) % (
            R['month'] + ('-preview' if pv else '')),
        '__ALERT_ROWS__': alert_rows(R, ('A', 'B', 'N'), empty_ab),
        '__C_DETAILS__': c_details,
        '__DORMANT_ROWS__': dormant_rows(R), '__ANOMALY_ROWS__': anomaly_rows(R),
        '__GROWTH_ROWS__': growth_rows(R), '__NEWRET_ROWS__': newret_rows(R),
        '__TOP_ROWS__': top_rows(R),
        '__ALL_ROWS__': all_rows(R), '__ALL_N__': str(len(R.get('all_items') or [])),
        '__DATA__': json.dumps(R, ensure_ascii=False).replace('</', '<\\/'),
    }
    out = TPL
    for k, v in repl.items():
        out = out.replace(k, v)
    return out


def build_index():
    formal, previews = {}, {}
    for fn in os.listdir(REPORTS):
        m = re.match(r'^(\d{4})-(\d{2})(-preview)?\.json$', fn)
        if not m:
            continue
        try:
            with open(os.path.join(REPORTS, fn), encoding='utf-8') as f:
                R = json.load(f)
        except Exception:
            continue
        (previews if m.group(3) else formal)[m.group(1) + '-' + m.group(2)] = R
    # 正式月報產生後，同月份的期中預覽即不再列出
    entries = [(k, R, False) for k, R in formal.items()]
    entries += [(k, R, True) for k, R in previews.items() if k not in formal]
    entries.sort(key=lambda e: e[0], reverse=True)
    rows = []
    for stem, R, is_pv in entries:
        t = R['totals']
        gc = R.get('grade_counts')
        if gc:
            alert_txt = '警示 A %d／B %d ｜ %s %d' % (
                gc['A'], gc['B'], '尚未生產' if is_pv else '斷單', len(R['dormant']))
        else:
            alert_txt = '警示 %d ｜ 斷單 %d' % (R['real_alert_count'], len(R['dormant']))
        month_txt = esc(R['month'])
        if is_pv:
            month_txt += '<div class="pv">期中預覽<br>截至 %s</div>' % esc(R['preview']['asof'][5:].replace('-', '/'))
        rows.append(
            '<a class="row%s" href="%s%s.html"><div class="m">%s</div>'
            '<div class="light %s">%s %s</div>'
            '<div class="n">%s 鍋　環比 %s　同比 %s</div>'
            '<div class="a">%s</div></a>' % (
                ' preview' if is_pv else '', stem, '-preview' if is_pv else '', month_txt,
                R['light'], LIGHT_ICON.get(R['light'], ''), esc(R['light_label']),
                fmt(t['pots']), pct_span(t['mom_pct']), pct_span(t['yoy_pct']), alert_txt))
    body = '\n'.join(rows) if rows else '<p class="muted">尚未產生任何月報。</p>'
    html_out = """<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>生產鍋數月報 · 索引</title>
<style>
body{margin:0;background:#f6f7f9;color:#0f172a;font-family:'Noto Sans TC',-apple-system,'Segoe UI',sans-serif;}
.wrap{max-width:820px;margin:0 auto;padding:34px 20px 60px;}
h1{font-size:26px;margin:0 0 6px;} .sub{font-size:15px;color:#64748b;margin-bottom:22px;line-height:1.6;}
.row{display:grid;grid-template-columns:110px 130px 1fr auto;gap:14px;align-items:center;
 background:#fff;border:1px solid #e2e8f0;border-radius:11px;padding:14px 17px;margin-bottom:10px;
 text-decoration:none;color:inherit;box-shadow:0 1px 3px rgba(15,23,42,.05);}
.row:hover{border-color:#1d4ed8;}
.row.preview{border:2px dashed #fb923c;background:#fffaf5;}
.pv{font-size:12.5px;font-weight:600;color:#c2410c;font-family:'Noto Sans TC',sans-serif;line-height:1.35;margin-top:3px;}
.m{font-family:ui-monospace,monospace;font-weight:700;font-size:18px;}
.light{font-size:13.5px;padding:4px 10px;border-radius:20px;text-align:center;white-space:nowrap;}
.light.red{background:#fee2e2;color:#991b1b;}.light.yellow{background:#fef3c7;color:#92400e;}.light.green{background:#dcfce7;color:#166534;}
.n{font-size:15px;color:#334155;font-family:ui-monospace,monospace;}
.a{font-size:13.5px;color:#64748b;white-space:nowrap;}
.up{color:#16a34a;}.dn{color:#dc2626;}.flat{color:#94a3b8;}
.muted{color:#94a3b8;font-size:13px;}
a.back{display:inline-block;margin-bottom:18px;font-size:14.5px;color:#1d4ed8;text-decoration:none;}
@media(max-width:640px){.row{grid-template-columns:1fr;gap:5px;}}
</style></head><body><div class="wrap">
<a class="back" href="../">← 回到儀表板</a>
<h1>生產鍋數月報</h1>
<div class="sub">依業務量（鍋數／重量）分析接單趨勢與衰退警示。點擊月份查看完整報告。橘色虛線框為「期中預覽」（本月尚未結束，與上月／去年同期比較）。</div>
__ROWS__
</div></body></html>"""
    return html_out.replace('__ROWS__', body)


def main():
    ap = argparse.ArgumentParser(description='產生月度業務分析報告')
    ap.add_argument('--month', required=True, help='目標月份，格式 YYYY-MM，例如 2026-08')
    ap.add_argument('--narrative', help='總評文字檔（純文字／段落以空行分隔，支援 **粗體**）')
    ap.add_argument('--xlsx', default=XLSX, help='資料來源 xlsx（預設 data/latest.xlsx）')
    ap.add_argument('--allow-incomplete', action='store_true', help='即使該月未補齊仍產生報告')
    ap.add_argument('--preview', action='store_true',
                    help='期中預覽：本月尚未結束，改與上月／去年「同期」比較，輸出 YYYY-MM-preview.html')
    ap.add_argument('--asof', help='期中預覽的資料截止日 YYYY-MM-DD（預設為資料中該月最後一天）')
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
    recs, meta, month_counts, _last = aggregate(rows)
    # 不納入分析的品項（item_overrides 的 status=excluded，例如資料有誤的代工品）：所有月份的紀錄一律剔除
    excluded = {c for c, o in (cfg.get('item_overrides') or {}).items() if o.get('status') == 'excluded'}
    recs = [r for r in recs if meta[r[0]]['料號'] not in excluded]

    if (Y, M) not in month_counts:
        have = sorted(month_counts)
        raise SystemExit('資料中沒有 %s 的紀錄。目前資料涵蓋 %s ~ %s。' % (
            mlabel(Y, M), mlabel(*have[0]), mlabel(*have[-1])))

    preview = None
    if args.preview:
        days = [d for (_k, ym, d, _p, _kg) in recs if ym == (Y, M)]
        if not days:
            raise SystemExit('%s 目前沒有有效的生產紀錄，無法產生預覽。' % mlabel(Y, M))
        D = max(days)
        if args.asof:
            a = parse_ymd(args.asof)
            if not a or (a[0], a[1]) != (Y, M):
                raise SystemExit('--asof 必須是 %s 內的日期，例如 %d-%02d-20' % (mlabel(Y, M), Y, M))
            D = a[2]
        # 截止日之後的紀錄一律不看，確保預覽內容與「當天看到的資料」一致
        recs = [r for r in recs if r[1] < (Y, M) or (r[1] == (Y, M) and r[2] <= D)]
        dim = calendar.monthrange(Y, M)[1]
        py_, pm_ = month_add(Y, M, -1)
        prev_dim = calendar.monthrange(py_, pm_)[1]
        full_month = D >= dim
        preview = {
            'asof': '%d-%02d-%02d' % (Y, M, D), 'day': D, 'days_in_month': dim,
            'full_month': full_month,
            'cur_period': '%d/1–%d/%d' % (M, M, D),
            'prev_period': '%d/1–%d/%d' % (pm_, pm_, prev_dim if full_month else min(D, prev_dim)),
            'yoy_period': '%d/%d/1–%d/%d' % (Y - 1, M, M, dim if full_month else D),
        }
        # 本月連續無生產紀錄的日期區間（≥ preview_gap_days 天）
        have = {d for (_k, ym, d, _p, _kg) in recs if ym == (Y, M)}
        gap_min = int(cfg.get('preview_gap_days', 5))
        gaps, run = [], 0
        for d in range(1, D + 2):
            if d <= D and d not in have:
                run += 1
                continue
            if run >= gap_min:
                g = {'range': '%d/%d–%d/%d' % (M, d - run, M, d - 1), 'days': run}
                # 已知的計畫性停產（report_config.json 的 planned_shutdowns）不當成資料漏登
                g0, g1 = date(Y, M, d - run), date(Y, M, d - 1)
                for s in cfg.get('planned_shutdowns') or []:
                    s0, s1 = parse_ymd(s.get('start')), parse_ymd(s.get('end'))
                    if s0 and s1 and date(*s0) <= g1 and date(*s1) >= g0:
                        g['planned'] = s.get('reason') or '計畫性停產'
                gaps.append(g)
            run = 0
        preview['gaps'] = gaps
        br = blank_rows(rows, Y, M, D, skip_codes=excluded)
        if br:
            preview['blank'] = {
                'rows': len(br),
                'semi': sum('半成品數' in b[3] for b in br),
                'fg': sum('成品數' in b[3].replace('半成品數', '') for b in br),
                'days': '、'.join('%d/%d' % (M, d) for d in sorted({b[0] for b in br})),
            }
        nk_days = sorted({d for (_k, ym, d, p, kv) in recs if ym == (Y, M) and p and p > 0 and not kv})
        if nk_days:
            shown = nk_days if len(nk_days) <= 8 else nk_days[-8:]
            preview['nokg_days'] = ('分布在 ' if len(nk_days) <= 8 else '主要在 ') + \
                '、'.join('%d/%d' % (M, d) for d in shown) + \
                ('（共 %d 天）' % len(nk_days) if len(nk_days) > 8 else '')
        full = roll(recs)
        cmp = full if full_month else roll(recs, D)
    else:
        full = roll(recs)
        cmp = full

    R = compute(full, cmp, meta, month_counts, Y, M, cfg, preview)
    if R['incomplete'] and not args.allow_incomplete and not preview:
        raise SystemExit(
            '%s 的資料可能尚未補齊（僅 %d 筆，約為其他月份中位數 %d 的 %d%%）。\n'
            '請確認資料完整後再產生；若是月底開會前要先看，請改用 --preview 產生期中預覽；\n'
            '若確定要產出正式版，加上 --allow-incomplete。' % (
                R['month'], R['record_count'], R['median_count'],
                round(R['record_count'] / R['median_count'] * 100) if R['median_count'] else 0))
    # 正式月報必須資料無空缺：生管尚未填完「半成品數(實際)」或「成品數(實際)」時只能出期中預覽
    if not preview and not args.allow_incomplete:
        br = blank_rows(rows, Y, M, skip_codes=excluded)
        if br:
            days = sorted({b[0] for b in br})
            raise SystemExit(
                '%s 的資料尚有空缺：%d 筆有鍋數，但半成品數（實際）%d 筆、成品數（實際）%d 筆未填，分布在 %s。\n'
                '正式月報需資料無空缺，請生管補齊後再產生；目前請改用 --preview 產生期中預覽。\n'
                '若確定要產出正式版，加上 --allow-incomplete。' % (
                    mlabel(Y, M), len(br),
                    sum('半成品數' in b[3] for b in br), sum('成品數' in b[3].replace('半成品數', '') for b in br),
                    '、'.join('%d/%d' % (M, d) for d in days)))

    narrative = None
    if args.narrative:
        with open(args.narrative, encoding='utf-8') as f:
            narrative = f.read()

    os.makedirs(REPORTS, exist_ok=True)
    stem = '%04d-%02d' % (Y, M) + ('-preview' if preview else '')
    with open(os.path.join(REPORTS, stem + '.json'), 'w', encoding='utf-8') as f:
        json.dump(R, f, ensure_ascii=False, indent=1)
    with open(os.path.join(REPORTS, stem + '.html'), 'w', encoding='utf-8') as f:
        f.write(build_html(R, narrative))
    with open(os.path.join(REPORTS, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(build_index())

    t = R['totals']
    gc = R['grade_counts']
    print('✓ 已產生 reports/%s.html' % stem)
    if preview:
        print('  期中預覽：資料截至 %s（與上月 %s、去年 %s 同期比較）' % (
            preview['asof'], preview['prev_period'], preview['yoy_period']))
    print('  總鍋數 %s（環比 %s／同比 %s）｜ 重量 %s kg ｜ 品項 %d' % (
        fmt(t['pots']), pct_txt(t['mom_pct']), pct_txt(t['yoy_pct']), fmt(t['kg']), t['items']))
    print('  燈號 %s %s ｜ 警示 A %d／B %d／C %d ｜ %s %d 項 ｜ 資料異常 %d 項' % (
        LIGHT_ICON.get(R['light'], ''), R['light_label'], gc['A'], gc['B'], gc['C'],
        '尚未生產' if preview else '斷單', len(R['dormant']), len(R['anomalies'])))


if __name__ == '__main__':
    main()
