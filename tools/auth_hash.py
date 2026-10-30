#!/usr/bin/env python3
"""產生 auth.js 帳號用的密碼雜湊值。

用法：
    python3 tools/auth_hash.py 帳號 密碼

把輸出的 hash 貼到 auth.js 的 ACCOUNTS 對應帳號。密碼明碼不要寫進任何檔案。
"""
import hashlib
import sys

SALT = 'pd-v1'  # 必須與 auth.js 的 SALT 相同


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    user, password = sys.argv[1], sys.argv[2]
    digest = hashlib.sha256(('%s|%s|%s' % (SALT, user, password)).encode('utf-8')).hexdigest()
    print("'%s': { name: '顯示名稱', admin: false, hash: '%s' }," % (user, digest))


if __name__ == '__main__':
    main()
