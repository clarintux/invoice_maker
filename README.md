<h1 align="center">
Invoice Maker
</h1>

<p align="center">
Cross-platform and multi-language automated invoice generator, wich support various invoice templates. Written in Python and wxPython.</a>.
</p>
<p align="center">
<img src="https://github.com/clarintux/invoice_maker/blob/master/icons/invoice_maker.jpg" height="400">
</p>
<hr />

## Preamble

A single generator that meets the different needs of everyone is probably impossible.
To try to achieve this, Invoice Maker lets the user choose the language, the appearance of invoices and other things.
Invoice Maker (v0.1) supports English, Italian, French, Spanish and German.
7 HTML templates are available.

## Installation

- On Linux:
  ```
  sudo apt install git wkhtmltopdf wxpython-tools python3-pdfkit  # requirements
  sudo ln -s /usr/bin/python3 /usr/bin/python  # symbolic link to use python3 with 'python' command
  git clone https://github.com/clarintux/invoice_maker.git  # clone this repository in '$HOME/invoice_maker'
  chmod +x ~/invoice_maker/invoice_maker.py   # permit execution
  chmod +x ~/invoice_maker/send_mail.py       # permit execution
  ```

- On Windows and Mac:
     - If not installed, install [Python](https://www.python.org/)
     - Install [wxPython](https://www.wxpython.org/):
     ```
     pip install wxPython
     ```
     - Install [wkhtmltopdf](https://wkhtmltopdf.org/):
     - Install [pdfkit](https://pypi.org/project/pdfkit/):
     ```
     pip install pdfkit
     ```
     - Clone the repository:
     ```
     git clone https://github.com/clarintux/invoice_maker.git
     ```


Replace the image `logo.jpg` with a JPG file with your logo.
Modify TEXT1, TEXT2 and TEXT3 in the file `text.txt` accordently to your needs.
Start invoice_maker and save your data and a invoice number.
Insert data and invoice items for each customer.

Enjoy!  :-)

P.S.: To automatically send invoices to customers, changes in the module `send_mail.py` are required. I can not predict wich port and server and things are used in your smtp server...

## Usage


- `python invoice_maker` - Launch the GUI
- `--help` - display help and exit
- `<locale code>` - use selected locale without GUI
- `--template` - choose a template between 1 and 7 without GUI
- `--all` - generate all invoices without GUI
- `--send` - send all generated invoices without GUI

## Examples

Assuming `invoice_maker.py` is in your path, it can be used like:
```
./invoice_maker.py
./invoice_maker.py --help
./invoice_maker.py it --all --template 1
./invoice_maker.py de --all --template 2
./invoice_maker.py en --all --template 3 --send
```

## Credits for Templates
- [https://codepen.io/tjoen/pen/wvgvLX](https://codepen.io/tjoen/pen/wvgvLX)
- [https://codepen.io/pvanfas/pen/vYEmGmK](https://codepen.io/pvanfas/pen/vYEmGmK)
- [https://bootsnipp.com/snippets/9gjD](https://bootsnipp.com/snippets/9gjD)
- [https://htmlpdfapi.com/blog/free_html5_invoice_templates](https://htmlpdfapi.com/blog/free_html5_invoice_templates)
- [https://sparksuite.github.io/simple-html-invoice-template/](https://sparksuite.github.io/simple-html-invoice-template/)
