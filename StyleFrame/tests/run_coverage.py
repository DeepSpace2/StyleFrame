import coverage
import webbrowser
from StyleFrame import tests

cov = coverage.Coverage(omit=[r'C:\Users\Adi\Documents\Python\StyleFrame\StyleFrame\warnings_conf.py',
                              r'C:\Users\Adi\Documents\Python\StyleFrame\StyleFrame\utils.py',
                              r'C:\Users\Adi\Documents\Python\StyleFrame\StyleFrame\tests\style_frame_tests.py'],
                        branch=True)
cov.exclude(r'raise')
cov.exclude(r'import')
cov.start()
tests.run()
cov.stop()
cov.save()
cov.html_report()
webbrowser.open_new(r'C:\Users\Adi\Documents\Python\StyleFrame\StyleFrame\tests\htmlcov\index.html')
