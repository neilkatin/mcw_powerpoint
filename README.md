
mcw_powerpoint
==============

This project tries to automate some repetative tasks
we do in the American Red Cross Mass Care Webinar Team.

The model is to add tags to our main powerpoint's notes section,
and then extract them to two files:

* a file that has announcements we will make in the Q&A panel of the webinar
  at the proper slide

* a speaker timings file that shows which speaker is doing which slide.
  This file allows us to track the webinar total time, and who should
  be on camera at each slide

Configuring
-----------

There are two configuration files: `config.py` and `.env`.

The major difference is: `.env` is not checked into git and should hold all security-sensitive values.

`config.py` can have arbitrary python statements.

These values must be defined for the program to work:

* `TENANT_ID`the red cross azure tenant UUID
* `LIENT_ID`the application id for this client in azure
* `LIENT_SECRET`the application secret.  Be sure to note the expiration time of the secret and add it as a comment in the .env file
* `RIVE_ID`the id of the Document Folder for the MCW team.

`TENANT_ID` and `DRIVE_ID` are probably not 'secret', but for now all the parameters are in the `.env` file

Running
----------

This project uses pipenv to manage dependencies and virtual environments.
There are several ways to install pipenv; I usually use the platform pip
to install a local per-user copy, and add that to my path.

The project should be relatively flexible as to python versions; any relatively
recent python3 version should work.  I tested on fedora 39 with python 3.12.

The rest of these instructions assume you are at the top of the git repo

``` shell
# install a virtual env; use your python version
PYTHON_VERSION="3.12"
pipenv --python "$PYTHON_VERSION"

# install required packages
pipenv --sync

# show valid arguments
pipenv run ./main.py --help

# do a default run with debug output
pipenv run ./main.py --debug

