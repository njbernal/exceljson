# Excel to JSON API

A simple API built on Flask and served using Waitress, dockerized and deployed to Google Cloud. <br />
Can be run locally in the docker, as a dev server using flask directly, or using Waitress.

The server does not store any data, it simple receives the FileObject from a **POST** request, parses it and returns the necessary **JSON** for all worksheets.

<br />

# Motivation

I'm sure this already exists out there, but as I do a *lot* of Excel to JSON conversions locally I decided to build my own. It is also an exercise in Python, and an item to add to my portfolio.

<br />

# Usage

## Parameters:
<ul>
<li><code>source</code>: *REQUIRED* .xlsx file to parse. Future support for .xls and .csv.</li>
<li><code>ignore_blank_rows</code>: *OPTIONAL* Set to true or 1, will ignore any fully blank rows, condensing documents.</li>
<li><code>limit</code>: *OPTIONAL* How many rows to parse and return.</li>
</ul>

<br />

## Remote use:
Send a **POST** request to https://exceljson-msgoe53wmq-wl.a.run.app/ <br />
Ensure the request is correctly formatted. 

<br />

## To use locally:
git clone https://github.com/njbernal/ExcelJSON.git
Create a Python virtualenv or install the dependencies in requirements.txt using pip.
Options to start the server:
1. Dockerize & run
2. Run <code>python server.py</code> from within the virtual environment.
3. Run <code>waitress-serve --listen=*:8888 server:app</code>