# Инструкция по запуску на Windows с использованием uv

## Требования
- Установленный [Python 3.10+](https://www.python.org/downloads/windows/)
- Установленный [uv](https://docs.astral.sh/uv/getting-started/installation/)

Через WinGet
```poweshell
winget install --id=astral-sh.uv  -e
```
Через irm
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```
## Запуск
По дефолту на вход принимается файл с именем `input.xlsx`, но можно задать свое имя как аргумент коммандной строки

default
```powerlshell
uv run main.py 
```
с передачей своего имени файла
```powershell
uv run main.py my_filename.xlsx
```

лист должен называться `school_marks_count`. Весь вывод идёт в папку `output`, название каждого файла - таймстамп.

