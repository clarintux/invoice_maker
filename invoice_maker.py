#!/usr/bin/env python

################################################################
#                                                              #
# Invoice Maker (v0.1): Cross-platform and multi-language      #
#                       automated invoce generator that        #
#                       supports various invoice templates.    #
#                                                              #
# Author: Clarintux                                            #
# Website: www.clarintuxmail.eu.org                            #
# Email: clarintux ( at) clarintuxmail (dot) eu (dot) org      #
# Copyright: (C) 2023                                          #
# License: GPL-3.0                                             #
################################################################

import os
import sys
import datetime
import time
import locale
import gettext
import re
import pickle
import pdfkit
#from weasyprint import HTML, CSS
from pathlib import Path
import wx
import wx.adv
import wx.grid
import wx.lib.agw.thumbnailctrl as TC
from send_mail import send


_ = wx.GetTranslation

appname = 'invoice_maker'
localedir = './locale'
LANGUAGES = ['en', 'it', 'de', 'es', 'fr'] # supported languages
NUM_COLS = 6 # 'quantity', 'article', 'description', 'date', 'price', 'total'
NUM_ROWS = 100 # max num articles
NUM_TEMPLATES = 7 # num supported templates
MAX_QUANTITY = 100 # max qunatity for each article

wantSend = False # send invoices to customers?
password = ""  # we prompt once for our email password and store it here


def changeContent2(contentTemplate, items, numItems):
    """Insert items in content and return content. Only for template2"""
    for i in range(numItems):
        contentTemplate = re.sub(r'(<!--ITEMS2-->)',
                                 r'\1<tr><td><a class="cut">-</a><span contenteditable>' + items[i]['article'] + '</span></td><td><span contenteditable>' +
                                 items[i]['description'] + ' ' + items[i]['date']+ '</span></td><td><span data-prefix>' + locale.localeconv()['currency_symbol'] +
                                 '</span><span contenteditable>' + items[i]['price'] + '</span></td><td><span contenteditable>' + items[i]['quantity'] +
                                 '</span></td><td><span data-prefix>' + locale.localeconv()['currency_symbol'] + '</span><span>' +
                                 str(float(items[i]['quantity']) * float(items[i]['price'])) + '</span></td></tr>', contentTemplate)

    return contentTemplate


def changeContent3(contentTemplate, items, numItems):
    """Insert items in content and return content. Only for template3"""
    for i in range(numItems):
        localPrice = locale.currency(float(items[i]['price']), grouping=True)
        contentTemplate = re.sub(r'(<!--ITEMS3-->)',
                                 r'\1<tr><td style="font-size: 12px; font-family: "Open Sans", sans-serif; color: #ff0000;  line-height: 18px;  vertical-align: top; padding:10px 0;" class="article">' +
                                 items[i]['article'] + '</td><td style="font-size: 12px; font-family: "Open Sans", sans-serif; color: #646a6e;  line-height: 18px;  vertical-align: top; padding:10px 0;"><small>' +
                                 items[i]['description'] + ' ' + items[i]['date'] + '</small></td><td style="font-size: 12px; font-family: "Open Sans", sans-serif; color: #646a6e;  line-height: 18px;  vertical-align: top; padding:10px 0;" align="center">'
                                 + items[i]['quantity'] + '</td><td style="font-size: 12px; font-family: "Open Sans", sans-serif; color: #1e2b33;  line-height: 18px;  vertical-align: top; padding:10px 0;" align="right">' +
                                 localPrice + '</td></tr><tr><td height="1" colspan="4" style="border-bottom:1px solid #e4e4e4"></td></tr>',
                                 contentTemplate)

    return contentTemplate


def changeContent4(contentTemplate, items, numItems, customerName, myMail):
    """Insert items in content and return content. Only for template4"""
    for i in range(numItems):
        localPrice = locale.currency(float(items[i]['price']), grouping=True)
        contentTemplate = re.sub(r'(<!--ITEMS4-->)',
                                 r'\1<tr class="item-row"><td class="item-name"><div class="delete-wpr"><textarea>' +
                                 items[i]['article'] + '</textarea><a class="delete" href="javascript:;" title="Remove row">X</a></div></td><td class="description"><textarea>' +
                                 items[i]['description'] + ' ' + items[i]['date'] + '</textarea></td><td><textarea class="cost">' +
                                 localPrice + '</textarea></td><td><textarea class="qty">' +
                                 items[i]['quantity'] + '</textarea></td><td><span class="price">' +
                                 locale.currency(float(items[i]['price']) * float(items[i]['quantity']), grouping=False) + '</span></td></tr>',
                                 contentTemplate)

    contentTemplate = contentTemplate.replace("<b>" + customerName + "</b>", customerName)
    contentTemplate = contentTemplate.replace('<a href = "mailto: ', '')
    contentTemplate = contentTemplate.replace('">' + myMail + '</a>', '')

    return contentTemplate


def changeContent5(contentTemplate, items, numItems):
    """Insert items in content and return content. Only for template5"""
    for i in range(numItems):
        localPrice = locale.currency(float(items[i]['price']), grouping=True)
        contentTemplate = re.sub(r'(<!--ITEMS5-->)',
                                 r'\1<tr><td class="service">' + items[i]['article'] + '</td><td class="desc">' +
                                 items[i]['description'] + ' ' + items[i]['date'] + '</td><td class="unit">' +
                                 localPrice + '</td><td class="qty">' + items[i]['quantity'] +
                                 '</td><td class="total">' + locale.currency(float(items[i]['price']) * float(items[i]['quantity']), grouping=False) +
                                 '</td></tr>',
                                 contentTemplate)

    return contentTemplate


def changeContent6(contentTemplate, items, numItems):
    """Insert items in content and return content. Only for template6"""
    for i in range(numItems):
        localPrice = locale.currency(float(items[i]['price']), grouping=True)
        contentTemplate = re.sub(r'(<!--ITEMS6-->)',
                                 r'\1<tr><td class="no">' + str(numItems-i) + '</td><td class="desc"><h3>' +
                                 items[i]['article'] + '</h3>' + items[i]['description'] + ' ' + items[i]['date'] + '</td><td class="unit">' +
                                 localPrice + '</td><td class="qty">' + items[i]['quantity'] +
                                 '</td><td class="total">' + locale.currency(float(items[i]['price']) * float(items[i]['quantity']), grouping=False)
                                 + '</td></tr>',
                                 contentTemplate)

    return contentTemplate


def changeContent7(contentTemplate, items, numItems):
    """Insert items in content and return content. Only for template7"""
    for i in range(numItems):
        localPrice = locale.currency(float(items[i]['price']), grouping=True)
        contentTemplate = re.sub(r'(<!--ITEMS7-->)',
                                 r'\1<tr><td class="service">' + items[i]['article'] +
                                 '</td><td class="desc">' + items[i]['description'] + ' ' + items[i]['date'] + '</td><td class="unit">' +
                                 localPrice + '</td><td class="qty">' + items[i]['quantity'] +
                                 '</td><td class="total">' + locale.currency(float(items[i]['price']) * float(items[i]['quantity']), grouping=False) +
                                 '</td></tr>',
                                 contentTemplate)

    return contentTemplate


def makeInvoice(customerName, templateNum):
    """
    Generate a single invoice for the specified customer and
    with the selected template. This function replace strings
    in the HTML template with your data, customer's data
    and the current date. After generation, invoiceNumber
    is incremented for the next invoice.
    Return: invoicePDFName
    """
    if not customerName:
        makeAllInvoices(templateNum)
        return

    currentDate = datetime.datetime.today().strftime('%x') # local date rapresentation
    year = datetime.datetime.today().strftime('%Y')
    month = datetime.datetime.today().strftime('%B') # Month as locale’s full name.

    # Load your data
    try:
        with open("./my_business/my_infos.pkl", "rb") as myInfoFile:
            myInfos = pickle.load(myInfoFile)
        myName = myInfos['name']
        myAddress = myInfos['address']
        myCity = myInfos['city']
        myPhone = myInfos['phone']
        myMail = myInfos['mail']
        myVAT = myInfos['vat']
        myCC = myInfos['cc']
        invoiceNumber = myInfos['number']
    except:
        sys.exit(_(u"ERROR: Couldn't read file './my_business/my_infos.pkl'  :-("))

    # Load customer's data and invoice
    customerFile = "./my_customers/" + customerName + ".pkl"
    try:
        with open(customerFile, 'rb') as file:
            customerInfos = pickle.load(file)
            invoice = pickle.load(file)
        customerAddress = customerInfos['address']
        customerCity = customerInfos['city']
        tax = customerInfos['tax']
    except:
        sys.exit(_(u"ERROR: Couldn't read file '" + customerFile + "'  :-("))

    total = invoice[0][NUM_COLS-1] # total is stored here (last column)
    localTotal = locale.currency(float(total), grouping=True)
    totalTax = float(tax) / 100.0 * float(total) + float(total)
    localTotalTax = locale.currency(totalTax, grouping=True)

    # template file
    templateFile = "./templates/template" + str(templateNum) + ".html"

    # Load template
    try:
        with open(templateFile, "r") as file:
            contentTemplate = file.read()
    except:
        sys.exit(_(u"ERROR: Couldn't read file '" + templateFile + "'  :-("))

    # Load additional text from file text.txt
    # 3 strings are mandatory here: the user
    # have to change the 3 default strings, if he wish.
    try:
        with open("./text.txt", "r") as file:
            line1 = file.readline()
            text1 = line1[line1.find("=")+1:]
            line2 = file.readline()
            text2 = line2[line2.find("=")+1:]
            line3 = file.readline()
            text3 = line3[line3.find("=")+1:]
    except:
        sys.exit(_(u"ERROR: Couldn't read file 'text.txt'  :-("))

    # Modify template with actual data
    contentTemplate = contentTemplate.replace("Hello", _(u"Hello"))
    contentTemplate = contentTemplate.replace("Thank you!", _(u"Thank you!"))
    contentTemplate = contentTemplate.replace("Thank you for your order", _(u"Thank you for your order"))
    contentTemplate = contentTemplate.replace("Invoice_", _(u"Invoice"))
    contentTemplate = contentTemplate.replace("INVOICE_TO_", _(u"INVOICE TO"))
    contentTemplate = contentTemplate.replace("CLIENT", _(u"CLIENT"))
    contentTemplate = contentTemplate.replace("COMPANY", _(u"COMPANY"))
    contentTemplate = contentTemplate.replace("ADDRESS", _(u"ADDRESS"))
    contentTemplate = contentTemplate.replace("Invoice #", _(u"Invoice Nr."))
    contentTemplate = contentTemplate.replace("invoice_number", invoiceNumber)
    contentTemplate = contentTemplate.replace("Date", _(u"Date"))
    contentTemplate = contentTemplate.replace("PhoneNumber", _(u"Phone"))
    contentTemplate = contentTemplate.replace("current_date", currentDate)
    contentTemplate = contentTemplate.replace("Amount Due", _(u"Amount Due"))
    contentTemplate = contentTemplate.replace("Amount Paid", _(u"Amount Paid"))
    contentTemplate = contentTemplate.replace("Balance Due", _(u"Balance Due"))
    contentTemplate = contentTemplate.replace("TAX_", _(u"TAX"))
    contentTemplate = contentTemplate.replace("currency_symbol", locale.localeconv()['currency_symbol'])
    contentTemplate = contentTemplate.replace("tot_amount", total)
    contentTemplate = contentTemplate.replace("Customer's name", "<b>" + customerName + "</b>")
    contentTemplate = contentTemplate.replace("Customer's address", customerAddress)
    contentTemplate = contentTemplate.replace("Customer's city", customerCity)
    contentTemplate = contentTemplate.replace("My Contact", _(u"Contact"))
    contentTemplate = contentTemplate.replace("My Name", myName)
    contentTemplate = contentTemplate.replace("My Address", myAddress)
    contentTemplate = contentTemplate.replace("My City", myCity)
    contentTemplate = contentTemplate.replace("My Phone", myPhone)
    contentTemplate = contentTemplate.replace("My Email", myMail)
    contentTemplate = contentTemplate.replace("Bank account", _("Bank account"))
    contentTemplate = contentTemplate.replace("MyBankAccount", myCC)
    contentTemplate = contentTemplate.replace("VAT Number", _("VAT Number"))
    contentTemplate = contentTemplate.replace("MyVatNumber", myVAT)
    contentTemplate = contentTemplate.replace("Item", _(u"Item"))
    contentTemplate = contentTemplate.replace("Service", _(u"Service"))
    contentTemplate = contentTemplate.replace("Description", _(u"Description"))
    contentTemplate = contentTemplate.replace("Price", _(u"Price"))
    contentTemplate = contentTemplate.replace("Quantity", _(u"Quantity"))
    contentTemplate = contentTemplate.replace("Total", _(u"Total"))
    contentTemplate = contentTemplate.replace("TEXT1", text1)
    contentTemplate = contentTemplate.replace("TEXT2", text2)
    contentTemplate = contentTemplate.replace("TEXT3", text3)
    contentTemplate = contentTemplate.replace("Additional Notes", _("Additional Notes"))
    contentTemplate = contentTemplate.replace("NOTICE", _("NOTICE"))
    contentTemplate = contentTemplate.replace("my_total" , localTotal)
    contentTemplate = contentTemplate.replace("my_tax" , tax)
    contentTemplate = contentTemplate.replace("MyTax" , localTotalTax)
    contentTemplate = contentTemplate.replace("plus_" , locale.format_string("%.2f", (float(tax) / 100.0 * float(total))))

    # make new list with only occupied rows of "invoice"
    numItems = 0
    items = []
    for row in range(NUM_ROWS):
        # if the row is empty check the next row.
        if not invoice[row][1] and not invoice[row][2] and not invoice[row][3] and not invoice[row][4]:
            continue
        else:
            items.append({"quantity": invoice[row][0],
                          "article": invoice[row][1],
                          "description": invoice[row][2],
                          "date": invoice[row][3],
                          "price": invoice[row][4]})
            numItems += 1

    if templateNum == 2:  # Prepare items for template2
        contentTemplate = changeContent2(contentTemplate, items, numItems)

    elif templateNum == 3: # Prepare items for template3
        contentTemplate = changeContent3(contentTemplate, items, numItems)

    elif templateNum == 4: # Prepare items for template4
        contentTemplate = changeContent4(contentTemplate, items, numItems, customerName, myMail)

    elif templateNum == 5:
        contentTemplate = changeContent5(contentTemplate, items, numItems)

    elif templateNum == 6:
        contentTemplate = changeContent6(contentTemplate, items, numItems)

    elif templateNum == 7:
        contentTemplate = changeContent7(contentTemplate, items, numItems)

    elif templateNum == 1:
        # Now Template1. A negativ value of numItems means invoice has no items (is empty)
        numItems -= 1
        if numItems >= 0:
            localPrice = locale.currency(float(items[numItems]['price']), grouping=True)
            # Prepare the last item in template1.
            if items[numItems]['quantity'] != '1':
                contentTemplate = contentTemplate.replace('<tr class="item last"></tr>', '<tr class="item last"><td>'
                                                          + items[numItems]['quantity'] + ' x ' + items[numItems]['article']
                                                          + ' ' + items[numItems]['description'] + ' ' + items[numItems]['date']
                                                          + '</td><td>' + localPrice + '</td>')
            else:
                contentTemplate = contentTemplate.replace('<tr class="item last"></tr>', '<tr class="item last"><td>'
                                                          + items[numItems]['article']
                                                          + ' ' + items[numItems]['description'] + ' ' + items[numItems]['date']
                                                          + '</td><td>' + localPrice + '</td>')


        # Prepare the others items in template1.
        numItems -= 1
        while (numItems >= 0):
            localPrice = locale.currency(float(items[numItems]['price']), grouping=True)

            if items[numItems]['quantity'] != '1':
                contentTemplate = re.sub(r'(<!--ITEMS-->)',
                                         r'\1<tr class="item"><td>' + items[numItems]['quantity'] + ' x ' + items[numItems]['article'] + ' ' + items[numItems]['description'] + ' ' + items[numItems]['date'] + '</td><td>' + localPrice + '</td></tr>',
                                         contentTemplate)
            else:
                contentTemplate = re.sub(r'(<!--ITEMS-->)',
                                         r'\1<tr class="item"><td>' + items[numItems]['article'] + ' ' + items[numItems]['description'] + ' ' + items[numItems]['date'] + '</td><td>' + localPrice + '</td></tr>',
                                         contentTemplate)

            numItems -= 1

    # If not exist, crete 'invoices' dir and subdirs of current year and month.
    # So we have invoices oraganised by month.
    path = "./invoices/" + year + "/" + month
    try:
        os.makedirs(path)
    except:
        pass

    # We wont the time in the filename, to not to overwrite old invoices in the same directory
    invoiceHtmlName = path + "/" + customerName + "_" + time.strftime('%Y_%m_%d_%H%M%S', time.localtime()) + ".html"
    try:
        with open(invoiceHtmlName, "w") as file:
            file.write(contentTemplate)
    except:
        sys.exit(_(u"ERROR: Couldn't write to file '" + invoiceHtmlName + "'  :-("))

    # We prepare a new invoiceNumber for the next automatically generated invoice
    invoiceNumber = int(invoiceNumber) + 1
    invoiceNumber = str(invoiceNumber)
    try:
        with open("./my_business/my_infos.pkl", "wb") as myInfoFile:
            myInfos = {'name': myName, 'address': myAddress, 'city': myCity, 'phone': myPhone, 'mail': myMail, 'vat': myVAT, 'cc': myCC, 'number': invoiceNumber}
            pickle.dump(myInfos, myInfoFile) # save the new invoiceNumber
    except:
        sys.exit(_(u"ERROR: Couldn't write to file './my_business/my_infos.pkl'  :-("))

    invoicePDFName = invoiceHtmlName.replace('.html', '.pdf')
    pdfkit.from_file(invoiceHtmlName, invoicePDFName, verbose=False, options={"enable-local-file-access": True})

    #html = HTML(invoiceHtmlName)
    #html.write_pdf(invoicePDFName)

    return invoicePDFName


def makeAllInvoices(templateNum):
    """
    Generate all invoices, invoking 'makeInvoice' for each customer
    in the 'my_customers' directory.
    """
    for customer in os.listdir('./my_customers'):
        if customer.endswith('.pkl'):
            customerName = customer.replace('.pkl', '')
            makeInvoice(customerName, templateNum)


def sendMails(templateNum):
    """
    Generate invoice for each customer and send email with PDF invoice
    as attachment. This function works by calling the 'send' function
    in the 'send_mail' module. Do not expect that it works without
    changing nothing in the 'send_mail' module. The user has to change
    variables in the module, according to his email server and liking.
    """
    global password # if we have a password, we don't need to prompt for it

    customerList = []
    customerCount = 0

    for customer in os.listdir('./my_customers'):
        if customer.endswith('.pkl'):
            customerName = customer.replace('.pkl', '')
            # generate invoice for each customer and store the filename of new invoice
            invoicePDFName = makeInvoice(customerName, templateNum)
            # Load customer's mail
            customerFile = "./my_customers/" + customerName + ".pkl"
            try:
                with open(customerFile, 'rb') as file:
                    customerInfos = pickle.load(file)
                customerMail = customerInfos['mail']
            except:
                sys.exit(_(u"ERROR: Couldn't read file '" + customerFile + "'  :-("))

            customerCount += 1
            customerList.append({'name': customerName, 'mail': customerMail, 'pdf': invoicePDFName})
    
    # If at lest one customer exist...send email
    if customerCount > 0:
        send(customerList, password)


class MainFrame(wx.Frame):
    """
    The main window

    Attributes
    ----------
    sMsg : StatusText (for statusbar)
    tcMyName : wx.TextCtrl
    tcCustomerName: wx.TextCtrl
    tcMyAddress : wx.TextCtrl
    tcCustomerAddress: wx.TextCtrl
    tcMyCity : wx.TextCtrl
    tcCustomerCity : wx.TextCtrl
    tcMyPhone : wx.TextCtrl
    tcMyMail : wx.TextCtrl
    tcCustomerMail : wx.TextCtrl
    tcMyVAT : wx.TextCtrl
    tcMyCC : wx.TextCtrl
    tcNumber : wx.TexCtrl
    tcTAX : wx.TexCtrl
    myGrid : wx.grid.Grid

    Methods
    -------
    createControls(self)
    connectControls(self)
    ResetGrid(self, evt)
    OnCellChange(self, evt)
    OnRowClick(self, evt)
    LoadMyInfos(self)
    onMenuRemoveClicked(self, evt)
    onMenuNewClicked(self, evt)
    onMenuOpenClicked(slf, evt)
    onMenuSaveClicked(self, evt)
    onMenuGenerateClicked(self, evt)
    onMenuGenAllClicked(self, evt)
    onMenuSendClicked(self, evt)
    onMenuExitClicked(self, evt)
    onMenuAboutClicked(self, evt)
    """
    def __init__(self):
        wx.Frame.__init__(self, None, title=wx.GetApp().GetAppName())
        self.SetIcon(wx.Icon('icons/invoice_icon.ico'))

        self.createControls()
        self.connectControls()

        self.sMsg = _(u"Ready") # Default message for the status bar
        self.SetStatusText(self.sMsg)

    def createControls(self):
        """
        Create statusbar, menubar, all text (StaticText and TextCtrl)
        and the grid
        """
        self.CreateStatusBar(1)

        # Create menubar
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        fileMenu.Append(wx.ID_NEW)
        fileMenu.Append(wx.ID_OPEN)
        fileMenu.Append(wx.ID_SAVE)
        fileMenu.Append(wx.ID_DELETE)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_EXIT)
        menubar.Append(fileMenu, wx.GetStockLabel(wx.ID_FILE))
        actionMenu = wx.Menu()
        actionMenu.Append(wx.ID_CLEAR)
        generateInvoiceMenu = wx.MenuItem(actionMenu, wx.ID_PRINT, _(u'&Generate invoice'))
        actionMenu.Append(generateInvoiceMenu)
        generateAllMenu = wx.MenuItem(actionMenu, wx.ID_SELECTALL, _(u'Generate &all invoices'))
        actionMenu.Append(generateAllMenu)
        sendMenu = wx.MenuItem(actionMenu, wx.ID_NETWORK, _(u'Generate and send &mails'))
        actionMenu.Append(sendMenu)
        menubar.Append(actionMenu, wx.GetStockLabel(wx.ID_PRINT))
        helpMenu = wx.Menu()
        helpMenu.Append(wx.ID_ABOUT)
        menubar.Append(helpMenu, wx.GetStockLabel(wx.ID_HELP))

        self.SetMenuBar(menubar)

        panel = wx.Panel(self, wx.ID_ANY)
        sizerMain = wx.BoxSizer(wx.VERTICAL)
        box1 = wx.GridBagSizer(5, 5)

        # Create text (StaticText and TextCtrl)
        textMyName = wx.StaticText(panel, label=_(u"Your name and surname*:"))
        box1.Add(textMyName, pos=(0, 0), flag=wx.TOP|wx.LEFT, border=5)
        tcMyName = wx.TextCtrl(panel)
        tcMyName.SetMinSize((240, 30))
        box1.Add(tcMyName, pos=(0, 1), flag=wx.TOP|wx.RIGHT, border=5)
        textcustomerName = wx.StaticText(panel, label=_(u"Customer's name*:"))
        box1.Add(textcustomerName, pos=(0, 3), flag=wx.TOP|wx.LEFT, border=5)
        tcCustomerName = wx.TextCtrl(panel)
        tcCustomerName.SetMinSize((240, 30))
        box1.Add(tcCustomerName, pos=(0, 4), flag=wx.TOP|wx.RIGHT, border=5)
        textMyAddress = wx.StaticText(panel, label=_(u"Your address (street and number)*:"))
        box1.Add(textMyAddress, pos=(1, 0), flag=wx.LEFT, border=5)
        tcMyAddress = wx.TextCtrl(panel)
        tcMyAddress.SetMinSize((240, 30))
        box1.Add(tcMyAddress, pos=(1, 1), flag=wx.RIGHT, border=5)
        textcustomerAddress = wx.StaticText(panel, label=_(u"Customer's address*:"))
        box1.Add(textcustomerAddress, pos=(1, 3), flag=wx.LEFT, border=5)
        tcCustomerAddress = wx.TextCtrl(panel)
        tcCustomerAddress.SetMinSize((240, 30))
        box1.Add(tcCustomerAddress, pos=(1, 4), flag=wx.RIGHT, border=5)
        textMyCity = wx.StaticText(panel, label=_(u"Your ZIP code and city*:"))
        box1.Add(textMyCity, pos=(2, 0), flag=wx.LEFT, border=5)
        tcMyCity = wx.TextCtrl(panel)
        tcMyCity.SetMinSize((240, 30))
        box1.Add(tcMyCity, pos=(2, 1), flag=wx.RIGHT, border=5)
        textcustomerCity = wx.StaticText(panel, label=_(u"Customer's ZIP code and city*:"))
        box1.Add(textcustomerCity, pos=(2, 3), flag=wx.LEFT, border=5)
        tcCustomerCity = wx.TextCtrl(panel)
        tcCustomerCity.SetMinSize((240, 30))
        box1.Add(tcCustomerCity, pos=(2, 4), flag=wx.RIGHT, border=5)
        textMyPhone = wx.StaticText(panel, label=_(u"Your phone number*:"))
        box1.Add(textMyPhone, pos=(3, 0), flag=wx.LEFT, border=5)
        tcMyPhone = wx.TextCtrl(panel)
        tcMyPhone.SetMinSize((240, 30))
        box1.Add(tcMyPhone, pos=(3, 1), flag=wx.RIGHT, border=5)
        textMyMail = wx.StaticText(panel, label=_(u"Your email address*:"))
        box1.Add(textMyMail, pos=(4, 0), flag=wx.LEFT, border=5)
        tcMyMail = wx.TextCtrl(panel)
        tcMyMail.SetMinSize((240, 30))
        box1.Add(tcMyMail, pos=(4, 1), flag=wx.RIGHT, border=5)
        textcustomerMail = wx.StaticText(panel, label=_(u"Customer's email address:"))
        box1.Add(textcustomerMail, pos=(4, 3), flag=wx.LEFT, border=5)
        tcCustomerMail = wx.TextCtrl(panel)
        tcCustomerMail.SetMinSize((240, 30))
        box1.Add(tcCustomerMail, pos=(4, 4), flag=wx.RIGHT, border=5)
        textMyVAT = wx.StaticText(panel, label=_(u"Your VAT number*:"))
        box1.Add(textMyVAT, pos=(5, 0), flag=wx.LEFT, border=5)
        tcMyVAT = wx.TextCtrl(panel)
        tcMyVAT.SetMinSize((240, 30))
        box1.Add(tcMyVAT, pos=(5, 1), flag=wx.RIGHT, border=5)
        textTAX = wx.StaticText(panel, label=_(u"TAX % to apply (0 means no tax)*:"))
        box1.Add(textTAX, pos=(5, 3), flag=wx.LEFT, border=5)
        tcTAX = wx.TextCtrl(panel)
        tcTAX.SetMinSize((240, 30))
        box1.Add(tcTAX, pos=(5, 4), flag=wx.RIGHT, border=5)
        textNumber = wx.StaticText(panel, label=_(u"Invoice number*:"))
        box1.Add(textNumber, pos=(6, 3), flag=wx.LEFT, border=5)
        tcNumber = wx.TextCtrl(panel)
        tcNumber.SetMinSize((240, 30))
        box1.Add(tcNumber, pos=(6, 4), flag=wx.RIGHT, border=5)
        textMyCC = wx.StaticText(panel, label=_(u"Your bank account*:"))
        box1.Add(textMyCC, pos=(6, 0), flag=wx.LEFT|wx.BOTTOM, border=5)
        tcMyCC = wx.TextCtrl(panel)
        tcMyCC.SetMinSize((240, 30))
        box1.Add(tcMyCC, pos=(6, 1), flag=wx.RIGHT|wx.BOTTOM, border=5)

        # set attributes
        self.tcMyName = tcMyName
        self.tcCustomerName = tcCustomerName
        self.tcMyAddress = tcMyAddress
        self.tcCustomerAddress = tcCustomerAddress
        self.tcMyCity = tcMyCity
        self.tcCustomerCity = tcCustomerCity
        self.tcMyPhone = tcMyPhone
        self.tcMyMail = tcMyMail
        self.tcCustomerMail = tcCustomerMail
        self.tcMyVAT = tcMyVAT
        self.tcMyCC = tcMyCC
        self.tcNumber = tcNumber
        self.tcTAX = tcTAX

        self.LoadMyInfos() # Load your data, if exist

        line = wx.StaticLine(panel)
        sizerMain.Add(box1)
        sizerMain.Add(line, 1, wx.EXPAND)

        # Create grid and set column labels and size
        myGrid = wx.grid.Grid(panel)
        self.myGrid = myGrid
        myGrid.CreateGrid(100, 6)
        myGrid.SetColLabelValue(0, _(u"Quantity"))
        myGrid.SetColLabelValue(1, _(u"Item or name"))
        myGrid.SetColLabelValue(2, _(u"Description (optional)"))
        myGrid.SetColLabelValue(3, _(u"Date (optional)"))
        myGrid.SetColLabelValue(4, _(u"Price"))
        myGrid.SetColLabelValue(5, _(u"Total"))
        myGrid.SetColSize(0 ,130)
        myGrid.SetColSize(1 ,225)
        myGrid.SetColSize(2 ,225)
        myGrid.SetColSize(3 ,120)
        myGrid.SetColSize(4 ,100)
        myGrid.SetColSize(5 ,100)
        
        self.ResetGrid(None)

        sizerMain.Add(myGrid, 1, wx.EXPAND)

        panel.SetSizer(sizerMain)
        self.Layout()
        self.SetMinSize((1070, 600))


    def connectControls(self):
        """Connect menu items to methods to perform actions"""
        self.Bind(wx.EVT_MENU, self.onMenuExitClicked, id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.onMenuAboutClicked, id=wx.ID_ABOUT)
        self.Bind(wx.EVT_MENU, self.onMenuNewClicked, id=wx.ID_NEW)
        self.Bind(wx.EVT_MENU, self.onMenuOpenClicked, id=wx.ID_OPEN)
        self.Bind(wx.EVT_MENU, self.onMenuSaveClicked, id=wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, self.onMenuRemoveClicked, id=wx.ID_DELETE)
        self.Bind(wx.EVT_MENU, self.ResetGrid, id=wx.ID_CLEAR)
        self.Bind(wx.EVT_MENU, self.onMenuGenerateClicked, id=wx.ID_PRINT)
        self.Bind(wx.EVT_MENU, self.onMenuGenAllClicked, id=wx.ID_SELECTALL)
        self.Bind(wx.EVT_MENU, self.onMenuSendClicked, id=wx.ID_NETWORK)
        self.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.OnCellChange)
        self.Bind(wx.grid.EVT_GRID_LABEL_LEFT_CLICK, self.OnRowClick)


    def ResetGrid(self, evt):
        """Clear cell values and set column formats"""
        self.myGrid.SetColFormatFloat(4, 6, 2) # 'price' column with float and precision 2
        self.myGrid.SetColFormatFloat(5, 6, 2) # 'total' column with float and precision 2
        for row in range(NUM_ROWS):
            self.myGrid.SetCellEditor(row, 0, wx.grid.GridCellNumberEditor(1,MAX_QUANTITY)) # 'quantity' column with int
            self.myGrid.SetCellValue(row, 0, "1") # default quantity = 1
            self.myGrid.SetReadOnly(row, 5) # 'total' column read only
            for column in range(1, NUM_COLS):
                self.myGrid.SetCellValue(row, column, "") # clear...
        self.myGrid.SetCellValue(0, 5, locale.str(0.00)) # default total = 0.0
        self.sMsg = _(u"Ready")
        self.SetStatusText(self.sMsg)


    def OnCellChange(self, evt):
        """Update total on cell change"""
        total = 0.0
        for row in range(NUM_ROWS):
            try:
                price = locale.atof(self.myGrid.GetCellValue(row, 4)) # 'price' column
                quantity = locale.atoi(self.myGrid.GetCellValue(row, 0)) # 'quantity' column
                total += price * quantity
            # If cell (row, 4) is empty, the try fails and we continue on the next row
            except ValueError:
                continue

        localTotal = locale.format_string("%.2f", float(total))
        self.myGrid.SetCellValue(0, 5, localTotal) # update total


    def OnRowClick(self, evt):
        """Delete a row by clicking on row headers"""
        row = evt.GetRow()

        # We don't wont the column headers...
        if row >= 0:
            dialog = wx.MessageDialog(self.GetParent(), _(u"Do you want to delete row number ") + str(row+1), _(u"Warning!"), wx.YES_NO| wx.NO_DEFAULT | wx.ICON_QUESTION | wx.STAY_ON_TOP)

            if dialog.ShowModal() == wx.ID_YES:
                # copy the next row until the last row
                for i in range(row, NUM_ROWS-1):
                    for column in range(NUM_COLS-1):
                        self.myGrid.SetCellValue(i, column, self.myGrid.GetCellValue(i+1, column))
                self.OnCellChange(None)
            dialog.Destroy()


    def LoadMyInfos(self):
        """Load your data if exist"""
        path = Path("./my_business/my_infos.pkl")
        if path.is_file():
            with open(path, 'rb') as file:
                myInfos = pickle.load(file)
            self.tcMyName.SetValue(myInfos['name'])
            self.tcMyAddress.SetValue(myInfos['address'])
            self.tcMyCity.SetValue(myInfos['city'])
            self.tcMyPhone.SetValue(myInfos['phone'])
            self.tcMyMail.SetValue(myInfos['mail'])
            self.tcMyVAT.SetValue(myInfos['vat'])
            self.tcMyCC.SetValue(myInfos['cc'])
            self.tcNumber.SetValue(myInfos['number'])
            return True
        else:
            return False


    def onMenuRemoveClicked(self, evt):
        """Delete customer's file: remove customer"""
        path = ""
        dialog = wx.FileDialog(self, _(u"Choose a invoice to remove..."), "", "", ".pkl file (*.pkl)|*pkl", wx.FD_OPEN)
        dialog.SetDirectory("./my_customers")
        if dialog.ShowModal() == wx.ID_OK:
            path = dialog.GetPath()
        dialog.Destroy()
        if path == "":
            return
        try:
            os.remove(path)
        except:
            return


    def onMenuNewClicked(self, evt):
        """Reset grid and customer's data on the window"""
        self.ResetGrid(None)
        self.tcCustomerName.SetValue("")
        self.tcCustomerAddress.SetValue("")
        self.tcCustomerCity.SetValue("")
        self.tcCustomerMail.SetValue("")
        self.tcTAX.SetValue("")
        self.LoadMyInfos()
        self.sMsg = _(u"Ready")
        self.SetStatusText(self.sMsg)


    def onMenuOpenClicked(self, evt):
        """Load customer's invoice"""
        path = ""
        dialog = wx.FileDialog(self, _(u"Choose a invoice to open..."), "", "", ".pkl file (*.pkl)|*pkl", wx.FD_OPEN)
        dialog.SetDirectory("./my_customers")
        if dialog.ShowModal() == wx.ID_OK:
            path = dialog.GetPath()
        dialog.Destroy()
        if path == "":
            return
        try:
            with open(path, 'rb') as file:
                customerInfos = pickle.load(file)
                invoice = pickle.load(file)
        except:
            self.sMsg = _(u"Error: Couln't load invoice file " + path + "  :-(")
            self.SetStatusText(self.sMsg)
            return

        self.tcCustomerName.SetValue(customerInfos['name'])
        self.tcCustomerAddress.SetValue(customerInfos['address'])
        self.tcCustomerCity.SetValue(customerInfos['city'])
        self.tcCustomerMail.SetValue(customerInfos['mail'])
        self.tcTAX.SetValue(customerInfos['tax'])

        for row in range(NUM_ROWS):
            for column in range(4): # 'quantity', 'article', 'description' and 'date' columns
                self.myGrid.SetCellValue(row, column, invoice[row][column])

        # load 'price' column and the total
        for row in range(NUM_ROWS):
            if invoice[row][4]:
                self.myGrid.SetCellValue(row, 4, locale.format_string("%.2f", float(invoice[row][4])))
        self.myGrid.SetCellValue(0, 5, locale.format_string("%.2f", float(invoice[0][5])))

        self.sMsg = _(u"Ready")
        self.SetStatusText(self.sMsg)


    def onMenuSaveClicked(self, evt):
        """
        Save your data (mandatory) or
        save your data and customer data (optional) or
        save your data and customer data and invoice (optional)
        """
        myName = self.tcMyName.GetValue()
        myAddress = self.tcMyAddress.GetValue()
        myCity = self.tcMyCity.GetValue()
        myPhone = self.tcMyPhone.GetValue()
        myMail = self.tcMyMail.GetValue()
        myVAT = self.tcMyVAT.GetValue()
        myCC = self.tcMyCC.GetValue()
        invoiceNumber = self.tcNumber.GetValue()
        customerName = self.tcCustomerName.GetValue()
        customerAddress = self.tcCustomerAddress.GetValue()
        customerCity = self.tcCustomerCity.GetValue()
        customerMail = self.tcCustomerMail.GetValue()
        tax = self.tcTAX.GetValue()

        # Return if the user don't insert mandatory infos of ihm self
        if not myName or not myAddress or not myCity or not myPhone or not myMail or not myVAT or not myCC or not invoiceNumber:
            self.sMsg = _(u"Error: First insert all your data and a invoice number!")
            self.SetStatusText(self.sMsg)
            return

        elif not invoiceNumber.isdigit():  # we need a integer to incremet invoice number value in the 'makeInvoice' function
            self.sMsg = _(u"Error: Invoice number must be a positiv integer!")
            self.SetStatusText(self.sMsg)
            return

        myInfos = {'name': myName, 'address': myAddress, 'city': myCity, 'phone': myPhone, 'mail': myMail, 'vat': myVAT, 'cc': myCC, 'number': invoiceNumber}

        # Save your data
        try:
            with open('./my_business/my_infos.pkl', 'wb') as file:
                pickle.dump(myInfos, file)
        except:
            self.sMsg = _(u"Error: Couln't write to file 'my_infos.pkl'  :-(")
            self.SetStatusText(self.sMsg)
            return

        self.sMsg = _(u"Saved your data and invoice number")

        # Save customer data if customername, customeraddress, customercity and tax exist
        if customerName and customerAddress and customerCity and tax:
            if not tax.isdigit() or int(tax) > 100:
                self.sMsg = _(u"Error: TAX must be a positiv integer less or equal than 100")
                self.SetStatusText(self.sMsg)
                return

            customerInfos = {'name': customerName, 'address': customerAddress, 'city': customerCity, 'mail': customerMail, 'tax': tax}
            invoice = [[column for column in range(6)] for row in range(100)] # invoice matrix
            filename = "./my_customers/" + customerName + ".pkl"
            for row in range(100):
                for column in range(NUM_COLS-1): # We don't need to copy 'total' column
                    invoice[row][column] = self.myGrid.GetCellValue(row, column)
            for row in range(NUM_ROWS):
                if invoice[row][4]:
                    invoice[row][4] = str(locale.atof(invoice[row][4])) # convert locale format to standard format
            invoice[0][5] = str(locale.atof(self.myGrid.GetCellValue(0, 5))) # convert locale format to standard format
            try:
                with open(filename, 'wb') as file:
                    pickle.dump(customerInfos, file)
                    pickle.dump(invoice, file)
            except:
                self.sMsg = _(u"Error: Couln't write to file " + filename + "  :-(")
                self.SetStatusText(self.sMsg)
                return
            
            self.sMsg = _(u"Saved your data and invoice for " + customerName)

        elif customerName or customerAddress or customerCity or customerMail: # if mandatory entries do not exist 
            self.sMsg = _(u"Error: First insert all customer's data!")
            self.SetStatusText(self.sMsg)
            return
        
        self.SetStatusText(self.sMsg)


    def onMenuGenerateClicked(self, evt):
        """Save and generate invoice for the specified customer"""
        customerName = self.tcCustomerName.GetValue()
        customerAddress = self.tcCustomerAddress.GetValue()
        customerCity = self.tcCustomerCity.GetValue()

        # Return if mandatory entries for the customer do not exist
        if not customerName or not customerAddress or not customerCity:
            self.sMsg = _(u"Cannot generate without customer's data...")
            self.SetStatusText(self.sMsg)
            return

        self.onMenuSaveClicked(None) # Save new data/invoice

        tFrame = TemplatesFrame(self, customerName) # new window to choose a invoice template

        self.sMsg = _(u"Generate invoice for ") + customerName
        self.SetStatusText(self.sMsg)


    def onMenuGenAllClicked(self, evt):
        """Save updates and generate invoices for all customers"""
        self.onMenuSaveClicked(None)
        
        customerCount = 0
        for customer in os.listdir('./my_customers'):
            if customer.endswith('.pkl'):
                customerCount += 1

        if customerCount == 0:
            self.sMsg = _(u"No customer's found...")
            self.SetStatusText(self.sMsg)
            return

        tFrame = TemplatesFrame(self, None) # new window to choose a invoice template
        self.sMsg = _(u"Generate all invoices")
        self.SetStatusText(self.sMsg)


    def onMenuSendClicked(self, evt):
        """Save updates and generate and send all invoices, if customer mail exist"""
        global wantSend
        self.onMenuSaveClicked(None)

        customerCount = 0
        for customer in os.listdir('./my_customers'):
            if customer.endswith('.pkl'):
                customerCount += 1

        if customerCount == 0:
            self.sMsg = _(u"No customer's founded...")
            self.SetStatusText(self.sMsg)
            return

        wantSend = True
        tFrame = TemplatesFrame(self, None) # new window to choose a invoice template


    def onMenuExitClicked(self, evt):
        """Quit the program by closing the window"""
        self.Close()


    def onMenuAboutClicked(self, evt):
        """Show About window"""
        info = wx.adv.AboutDialogInfo()
        info.SetName('Invoice Maker')
        info.SetVersion('(v0.1)')
        info.SetDescription(_(u'Automated invoices generator with various templates'))
        info.SetCopyright('(C) 2023')
        try:
            with open("LICENSE", "r") as licenseFile:
                licenseText = licenseFile.read()
        except:
            licenseText = 'GNU GENERAL PUBLIC LICENSE v3.0'
        info.SetLicense(licenseText)
        info.SetWebSite('http://clarintuxmail.eu.org')
        info.AddDeveloper('Clarintux\nclarintux@clarintuxmail.eu.org')
        info.AddTranslator('Clarintux')
        info.SetIcon(wx.Icon('icons/invoice_icon.png', wx.BITMAP_TYPE_PNG))
        wx.adv.AboutBox(info)        


class TemplatesFrame(wx.Frame):
    """
    Window to show thumbnails of the various templates
    and to choose one
    """
    def __init__(self, parent, customerName):
        self.customerName = customerName
        wx.Frame.__init__(self, parent, -1, _(u"Choose template..."), size=(580, 360))
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.thumbnail = TC.ThumbnailCtrl(panel, imagehandler=TC.NativeImageHandler)
        sizer.Add(self.thumbnail, 1, wx.EXPAND | wx.ALL, 10)
        self.thumbnail.ShowDir("./templates/thumbnails")
        self.thumbnail.SetThumbSize(144, 120)
        panel.SetSizer(sizer)
        self.Layout()
        self.SetMinSize((580, 360))
        self.SetWindowStyle(wx.STAY_ON_TOP)
        self.Show()
        try:
            # EVT_THUMBNAILS_DCLICK don't work in my system. That is a workaround
            self.thumbnail.Bind(TC.EVT_THUMBNAILS_DCLICK, self.SelectedThumb)
        except:
            self.thumbnail.Bind(TC.EVT_THUMBNAILS_CHAR, self.SelectedThumb)


    def SelectedThumb(self, evt):
        """Call makeInvoice with selected template"""
        global wantSend
        global password
        selection = self.thumbnail.GetSelection()
        selection += 1 # selection is 0 for template1, so we correct this 
        self.Close()
        if wantSend == True:
            if not password:
                dialog = wx.PasswordEntryDialog(self, _(u"Enter your email password"))
                retValue = dialog.ShowModal()
                if retValue == wx.ID_OK:
                    password = dialog.GetValue()
                else:
                    dialog.Close(True)
                dialog.Destroy()
            sendMails(selection)
        else:
            makeInvoice(self.customerName, selection)
        if wantSend == True:
            wx.MessageBox(_(u"Sended invoice(s)  :-)"), _("INFO"), wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(_(u"Generated invoice(s)  :-)"), _("INFO"), wx.OK | wx.ICON_INFORMATION)
        wantSend = False


class MyApp(wx.App):
    """Select language and init main window"""
    locale = None

    def OnInit(self):
        # Set Current directory to the one containing this file
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

        self.SetAppName('Invoice Maker')
        self.InitLanguage()

        try:
            localeCode = self.locale.GetName() + '.utf8' # Set locale to selected locale
        except:
            localeCode = 'en_US.utf8'  # default locale is english, if try fails
        
        locale.setlocale(locale.LC_ALL, localeCode)

        if localeCode != 'en_US.utf8':
            translations = gettext.translation(appname, localedir, fallback=True, languages=[self.locale.GetName()[slice(2)]])
            translations.install()

        # Create the main window
        frame = MainFrame()
        self.SetTopWindow(frame)
        frame.Show()
        return True


    def InitLanguage(self):
        """Select language"""
        langsAvail = {
                "English": wx.LANGUAGE_ENGLISH,
                "Italian": wx.LANGUAGE_ITALIAN,
                "German": wx.LANGUAGE_GERMAN,
                "Spanish": wx.LANGUAGE_SPANISH,
                "French": wx.LANGUAGE_FRENCH,
        }

        sel = wx.GetSingleChoice("Please choose language:", "Language", list(langsAvail.keys()))
        
        try:
            lang = langsAvail[sel]
        except:
            os._exit(1)

        # We can return, if english
        if lang == wx.LANGUAGE_ENGLISH:
            return

        wx.Locale.AddCatalogLookupPathPrefix("locale")
        self.locale = wx.Locale()
        if not self.locale.Init(lang):
            wx.LogWarning("This language is not supported by the system.")
        if not self.locale.AddCatalog("invoice_maker"):
            wx.LogError("Couldn't find/load the 'Invoice Maker' catalog for locale '" + self.locale.GetCanonicalName() + "'.")


if __name__ == '__main__':
    if len(sys.argv) > 1:
        try:
            localeCode = sys.argv[1] # locale code: for example <it>
            arg2 = sys.argv[2] # --all
            arg3 = sys.argv[3] # --template
            templateNum = sys.argv[4] # template num (1 - NUM_TEMPLATES)
        except:
            raise SystemExit(f"Usage: {sys.argv[0]} (without arguments)\n" +
                             f"   or: {sys.argv[0]} <locale code> --all --template <N>\n" +
                             f"   or: {sys.argv[0]} --help\n\n" +
                             "With no arguments, launch GUI.\n\n" +
                             "Arguments:\n" +
                                  "--all              generate all invoices\n" +
                                  "--send             generate all invoices and send to customers\n" +
                                  "--template <N>     choose template. Between 1 and " + str(NUM_TEMPLATES) + "\n" +
                                  "--help             display this help and exit\n\n" +
                            "Examples:\n" +
                                  f"{sys.argv[0]} it --all --template 1\n" +
                                  f"{sys.argv[0]} de --all --template 2 --send\n" +
                                  f"{sys.argv[0]} en --all --template 3\n\n" +
                             f"Supported languages: {LANGUAGES}")

        if localeCode not in LANGUAGES:
            print("'" + localeCode + "' is not a supported locale. Choose one of the following:", end=" ")
            for lang in LANGUAGES:
                print(lang, end=' ')
            sys.exit("")
        if arg2 == "--all" and arg3 == "--template":
            if int(templateNum) < 1 or int(templateNum) > NUM_TEMPLATES:
                sys.exit("Template must be between 1 and " + str(NUM_TEMPLATES))

            if localeCode  == "en":
                utf8 = "en_US.utf8"
            else:
                utf8 = localeCode + '_' + localeCode.upper() + '.utf8'
            locale.setlocale(locale.LC_ALL, utf8)
            translations = gettext.translation(appname, localedir, fallback=True, languages=[localeCode.strip()])
            translations.install()

            _ = translations.gettext
            
            try:
                if sys.argv[5] == "--send":
                    sendMails(int(templateNum))
            except:
                makeAllInvoices(int(templateNum))

            os._exit(0)

    app = MyApp()
    app.MainLoop()
