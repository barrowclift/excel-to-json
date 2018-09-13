Workbook to JSON
================

A small script I needed to write for a quick "one-off" project at a previous employer. Hopefully this will be useful to someone else out there as well.

Please note this is almost certainly not as flexible as it could or should be; it does not format Excel dates into a human-readable format, all number fields are doubles when saved (they always have a trailing ".0" afterwards), and boolean fields are saved with a 0 for False and 1 for True, to name a few rough edges. I'm sure there are other fields that aren't *quite* formatted as expected in the resulting JSON, but these are the big ones I know about.

Dependencies
------------

Requires Python 3 and [xlrd](https://pypi.python.org/pypi/xlrd) for raw Excel workbook parsing.

Usage
-----

	USAGE: xmlxToJson.py [OPTIONS] ... [FILE.xlsx|FILE.xls] ...
	    -p, --pretty
	        Pretty-print the built JSON files (off by default to help save space,
	        the generated JSON files are big)

	    -s N, --split=N
	        Split up the generated JSON into multiple files. For example, an N of
	        1000 would start a new JSON file when the 1000th root child element was
	        written.

	        The generated files will have the naming convention FILENAME_1.json,
	        FILENAME_2.json, etc.

	    -h, --help
	        Print the usage documentation
