<h1 align="center">
Invoice Maker
</h1>

<p align="center">
Cross-platform and multi-language automated invoice generator, wich support various invoice templates. Written in Python and wxPython.</a>.
</p>
<p align="center">
<img src="https://github.com/clarintux/invoice_maker/blob/master/icons/invoice_icon.png" height="400">
</p>
<hr />

## Preamble

A single generator that meets the different needs of everyone is probably impossible.
To try to achieve this, Invoice Maker lets the user choose the language, the appearance of invoices and other things.
Invoice Maker (v0.1) supports English, Italian, French, Spanish and German.
7 HTML templates are available.

## Installation

- Install [Python](https://www.python.org/), if not preinstalled in your System.
- Prerequisites for Linux:
* dpkg-dev
* build-essential
* python3-dev
* freeglut3-dev
* libgl1-mesa-dev
* libglu1-mesa-dev
* libgstreamer-plugins-base1.0-dev
* libgtk-3-dev
* libjpeg-dev
* libnotify-dev
* libpng-dev
* libsdl2-dev
* libsm-dev
* libtiff-dev
* libwebkit2gtk-4.0-dev
* libxtst-dev
- Install [wxPython](https://www.wxpython.org/):
```
pip3 install wxpython
```
- Install [wkhtmltopdf](https://wkhtmltopdf.org/):
```
sudo apt install wkhtmltopdf # (for debian-based systems)
```
- Install pdfkit:
```
pip3 install pdfkit
```

Replace the image `logo.jpg` with a JPG file with your logo.
Modify TEXT1, TEXT2 and TEXT3 in the file `text.txt` accordently to your needs.
Start invoice_maker and save your data and a invoice number.
Insert data and invoice items for each customer.
Enjoy!  :-)

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

./invoice_maker.py it --all --template 0
./invoice_maker.py de --all --template 1
./invoice_maker.py en --all --template 2 --send
```
