CS 4675 Summer 2026
Problem 1.1
Option: No AI

Using Apache Nutch as webcrawler
Using Apache Solr to index and search

Description:
Using Apache Nutch, this python file uses pandas to take the dump file created from apache nutch,
put it in a csv file and run data analysis on it.

Installations:
Apache Nutch
Apache Solr
see requirements.txt

To run:
set up apache nutch
Set java home in the terminal via
export JAVA_HOME = PATH

set seed url
(for my example I used cc.gatech.edu as the seed url)

run the crawl program and grab the dump file using readdb and -dump

put the path of this file into the DUMP_FILE variable.

Running this should create an output graph and output xslx, if you need to change in the constants in organize.py

also make sure hadloop.log ONLY includes the crawl iterations and not before and after because it will mess with the
crawl speed calculation.

Further analysis is found in the report file and in the output file that has keyword frequency, raw data, and graphs.
