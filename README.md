# json2excel

Simple command line tool to flatten JSON or JSON-like files and save them as csv or xlsx.

## Prerequisites

Writen and tested in Python 3.7.2.

## Installing

Clone the repository:
```
git clone https://github.com/tomasfarias/json2excel.git
```
Like usual, to install requirements simply run:
```
python -m pip install -r requirements.txt
```

## Running tests

Test file is located in tests directory and written using `pytest`. Add your current directory to PYTHONPATH first:

In Windows:
```
set PYTHONPATH=%cd%
```

In bash:
```
export PYTHONPATH=$pwd
```

Then run the tests:

```
pytest tests
```

## Author

* Tomás Farías - _All the work although little_ - [tomasfarias](https://github.com/tomasfarias)

## License

Standard MIT License.