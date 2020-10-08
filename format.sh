isort $(find pyexcel_xlsxw -name "*.py"|xargs echo) $(find tests -name "*.py"|xargs echo)
black -l 79 pyexcel_xlsxw
black -l 79 tests
