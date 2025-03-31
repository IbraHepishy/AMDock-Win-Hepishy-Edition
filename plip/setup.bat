@echo off
python -m venv env
call env\Scripts\activate.bat
python -m pip install openbabel-wheel==3.1.1.21
python setup.py install
python install_pymol.py
plip
pause