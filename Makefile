test:
	python -m doctest -v RemoveAttachments.py

check:
	python pep8.py RemoveAttachments.py
	pylint --rcfile pylintrc RemoveAttachments.py

install:
	python setup.py build
	python setup.py install

