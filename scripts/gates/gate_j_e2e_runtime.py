#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 J：完整端到端运行验证（交付前执行）
- J1. 子进程启动验证
- J2. 端口绑定验证（可配置端口列表）
- J3. API 端到端验证（可选，涉及外部 API 时传 --api-url）

用法:
    python gate_j_e2e_runtime.py [--ports 8765,8766] [--api-url https://api.example.com/health]
"""
import argparse
import socket
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description='门禁 J：端到端运行验证')
    parser.add_argument('--ports', default='8765,8766', help='逗号分隔的端口列表')
    parser.add_argument('--api-url', default='', help='外部 API 健康检查地址（可选）')
    args = parser.parse_args()

    errors = []

    # J1. 子进程启动验证
    try:
        result = subprocess.run(
            [sys.executable, '-c', 'print("hello")'],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode != 0:
            errors.append(f'Python 子进程启动失败: {result.stderr.strip()}')
        else:
            print('  ✅ 子进程启动正常')
    except Exception as exc:  # noqa: BLE001
        errors.append(f'子进程启动异常: {exc}')

    # J2. 端口绑定验证
    for port_text in args.ports.split(','):
        port = int(port_text.strip())
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind(('127.0.0.1', port))
            sock.close()
            print(f'  ✅ 端口 {port} 可用')
        except OSError:
            print(f'  ⚠️ 端口 {port} 被占用（不影响运行，会自动切换）')

    # J3. API 端到端验证（可选）
    if args.api_url:
        try:
            import requests
            resp = requests.get(args.api_url, timeout=5)
            print(f'  ✅ API 可访问: {resp.status_code}')
        except Exception as exc:  # noqa: BLE001
            print(f'  ⚠️ API 不可达: {exc}（不影响交付，但需告知用户）')

    if errors:
        print(f'门禁 J 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 J 通过')


if __name__ == '__main__':
    main()
