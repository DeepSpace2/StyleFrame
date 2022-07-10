### Thanks for contributing!

# Branch
Create the pull request against the `devel` branch instead of the `master`.

# Quality
All pull requests are welcome regardless of quality. We will work together to review and improve the pull request.

Of course there are some steps that could be considered to improve the quality of the pull request and to support the process of review and development.

- Make sure your code follows [PEP 8 — Style Guide for Python Code](https://pep8.org/). You can use some [linting tools](https://en.wikipedia.org/wiki/Lint_(software)) to automate the process (e.g. [`pycodestyle`](https://pycodestyle.pycqa.org) or [`flake8`](https://github.com/pycqa/flake8)).
- Test your code using the existing test cases. For details read further.
- It would be great if you could create [`unittest`](https://docs.python.org/3/library/unittest.html)'s for your code.

# Unittesting

This project uses pythons default [`unittest`](https://docs.python.org/3/library/unittest.html) package. You can run all test cases via the *discover* feature when you run this in the projects root folder.

```sh
python3 -m unittest discover
```

The output should look like this
```sh
----------------------------------------------------------------------
Ran 74 tests in 0.823s

OK
```

If you don't want to run all tests at once but a specific test case or method you need to *install* the package first. It is possible via a virtual environment but we recommend to use the `--editable` switch of `pip`. Run this in the projects root folder:
```sh
python3 -m pip install --editable .
```
Please read further to understand the consequences of `--editable`.
- ["pip documentation - Local project installs - Editable installs"](https://pip.pypa.io/en/stable/topics/local-project-installs/#editable-installs)
- ["When would the -e, --editable option be useful with pip install?"](https://stackoverflow.com/q/35064426/4865723)

### Happy coding!
