.. pyTNMR documentation master file, created by
   sphinx-quickstart on Fri Oct 24 17:00:58 2025.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pyTNMR documentation
====================

pyTNMR interfaces Python with Tecmag Inc.'s TNMR software to provide a significant amount of automation. When used in conjunction with Spyctra (https://github.com/spyctra/spyctra), active feedback of experimental results can be obtained to continuously audit data as it is acquired, change the observation frequency of a measurement, automatically repeat measurements, or actively interface with other experimental equipment, to name but a few.


.. py:function:: pyTNMR.__init__(self, path, unique=True, running=1)

    Initialize a pyTNMR object. Creates a subdirectory in the current working directory with name *path*. Copies the source code of the script into this directory and creates unique pyTNMR and experimental log files to help track issues or relevant variables.

    :param path: The name of the experimental directory, relative to the current working directory, where the .tnt files, logs, and source code are stored.
    :type path: str
    :param unique: A flag indicating if the experimental directory must be unique. If unique then a numerical index will be appended to *path* each time the code is rerun.
    :type unique: int[0 1] or bool
    :param running: A flag indicating if the sequences should actually be run (i.e. ZG performed). This can help with debugging longer, more complicated scripts.
    :type running: int[0, 1] or bool.


.. py:function:: pyTNMR.get_param(self, param)

    Return the value of the desired parameter.

    :param param: The parameter whose value is to be returned.
    :type param: str


.. py:function:: pyTNMR.get_table(self, table)

    Return the values of the desired table.

    :param table: The table whose value is to be returned.
    :type table: str


.. py:function:: pyTNMR.log(self, line)

    Print the desired line to the experimental log file in the experiment directory.

    :param line: The string being printed to the experimental log.
    :type line: str


.. py:function:: pyTNMR.open(self, filename)

    Open the requested tnt file. Will check experimental directory unless a more detailed path to the target .tnt file is specified.

    :param filename: The path of the file to open.
    :type filename: str


.. py:function:: pyTNMR.read(self, filename)

    Generate a spyctra object from the target .tnt file.

    :param filename: The path of the file to generate a .tnt from. Path is relative to experimental directory
    :type filename: str


.. py:function::  pyTNMR.reset(self)

    Resets the TNRM hardware (equivalent to Hammer button).


.. py:function:: pyTNMR.save_as(self, filename)

    Save the current active file as *filename*.

    :param filename: The name of the new .tnt file.
    :type filename: str


.. py:function:: pyTNMR.set_param(self, param, value)

    Set the sequence parameter to the desired *value*.

    :param param: Which parameter to change.
    :type param: str
    :param value: The new value of the target parameter.
    :type value:


.. py:function:: pyTNMR.set_table(self, table, value)

    Set the sequence table to the desired *value*.

    :param table: Which parameter to change.
    :type table: str
    :param value: The new value of the target parameter.
    :type value:


.. py:function:: pyTNMR.sleep(self, t)

    Stop the sequence for *t* seconds. Only works if *self.running* is True to help with debugging scripts.

    :param t: The amount of seconds to pause the script.
    :type t: numeric



.. py:function:: pyTNMR.zg(self, *manual_check)

    Run the sequence (equivalent to ZG button). Only works if *self.running* is True to help with debugging scripts. Will estimate time to completion and check that progress is as expected.

    :param manual_check: Flag to query user after each complete ZG to redo the experiment. Useful when manually interacting with experiment.
    :type manual_check: int[0,1] or bool




.. toctree::
   :maxdepth: 2
   :caption: Contents:

