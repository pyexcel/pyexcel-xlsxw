
pip freeze
nosetests --with-cov --cover-package pyexcel_xlsxw --cover-package tests --with-doctest --doctest-extension=.rst tests README.rst pyexcel_xlsxw  && flake8 . --exclude=.moban.d --builtins=unicode,xrange,long
