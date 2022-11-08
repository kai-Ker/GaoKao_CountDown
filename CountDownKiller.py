import os
import psutil

__author__ = "kai_Ker (kai_Ker@buaa.edu.cn)"

if __name__ == '__main__':
    try:
        with open('proc', 'r') as f:
            pid, name = f.read().split('|')
        proc = psutil.Process(int(pid))
        if proc.name().lower() == name.lower():
            proc.terminate()
            print(f'已结束进程: pid: {pid}, name: {name}')
            print('请手动刷新桌面')
        else:
            print(f'pid-{pid}的进程名为{proc.name()}, 而不是{name}, 可能程序已经退出, 请手动刷新桌面, 或者 proc 记录文件被错误修改')
    except FileNotFoundError:
        print('找不到 proc 记录文件, 可能你已经将其删除, 或更改了程序目录')
    except psutil.NoSuchProcess:
        print(f'没有进程pid为{pid}, 可能程序已经退出, 请手动刷新桌面')

    os.system('pause>nul')
