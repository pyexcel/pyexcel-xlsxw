{%extends 'README.rst.jj2' %}

{%block description%}
**{{name}}** is a tiny wrapper library to write data in xlsx and xlsm fromat using xlsxwriter. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`__.
{%endblock%}


{%block pagination%}
{%endblock%}

{%block read_from_file %}

Here's the sample code to help you read the data back. You will need to install pyexcel-xls or pyexcel-xlsx.

.. code-block:: python

    >>> from pyexcel_io import get_data
    >>> data = get_data("your_file.{{file_type}}")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

{%endblock%}

{%block read_from_memory %}

Here's the sample code to help you read the data back. You will need to install pyexcel-xls or pyexcel-xlsx.

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with {{file_type}} file upload
    >>> # where you will read from requests.FILES['YOUR_{{file_type|upper}}_FILE']
    >>> data = get_data(io, 'xlsx')
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}

{%endblock%}

{%block read_from_file_via_pyexcel %}

Let's assume we have data as the following.

.. code-block:: python

    >>> import pyexcel as pe
    >>> sheet = pe.get_book(file_name="your_file.{{file_type}}")
    >>> sheet
    Sheet 1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet 2:
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+

{%endblock%}

{%block read_from_memory_via_pyexcel %}
{%endblock%}
