pyinstaller -F -w --add-data "./ico.ico;ico.ico" main.py
Zx

python -m nuitka --standalone --windows-disable-console --mingw64 --nofollow-imports --show-progress --enable-plugin=tk-inter --windows-icon-from-ico=ico.ico --onefile main.py