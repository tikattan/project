#!-*-encoding=utf-8-*-
#!/usr/bin/env python

# Note merchant & ecommerce are from localhost database. Don't forget to run local server
# python import_orders.py <file> <ecommerce_code> <merchant_id> <optional *rows of line of order start with line 3>
# python import_orders.py be.xlsx BEX BE004 2 3 4

import os
import sys
import urllib2
import base64

from django.db import transaction
from branded_express.models import Order,Product

from xlrd import open_workbook
from mmap import mmap,ACCESS_READ
from django.db import transaction
from branded_express import models

import json

INCLUDE_HEADER=['index 0','BE LOGO','HEADER']
DEBUG = True
ORDER_BATCH_FILE_HEADERS=20
USERNAME='beserve'

# logistics site password
PASSWORD='4winter5'
# local password
PASSWORD='beserve'

URL='http://localhost:8000/api/orders/'
# URL='http://logistics.apitrans.com/api/orders/'

def log(msg, *args):
    sys.stdout.write((args and (msg % args) or msg))

def _get_basic_http_auth(username, password):
        auth = '%s:%s' % (username, password)
        auth = 'Basic %s' % base64.encodestring(auth)
        return auth.strip()

def send_request(_order):
    req = urllib2.Request(URL)
    req.add_header('Content-Type', 'application/json; charset=utf-8')
    req.add_header('AUTHORIZATION',_get_basic_http_auth(USERNAME,PASSWORD))
    print json.dumps(order, sort_keys=True,indent=4, separators=(',', ': '))
    response = urllib2.urlopen(req, json.dumps(_order, sort_keys=True,indent=4, separators=(',', ': ')))

if __name__ == "__main__":

    file = open(sys.argv[1], 'r')
    reader = open_workbook(file_contents=mmap(file.fileno(), 0, access=ACCESS_READ,))
    sheet = reader.sheet_by_index(0)

    if len(sys.argv) > 4:# [0,1,2,3]
        for number,data in enumerate(sys.argv[4:]):
            try:
                if int(data)<3:
                    sys.stderr.write("This command doesn't have parameter enough for orders")
                    sys.exit(1)
                elif int(data)>sheet.nrows:
                    sys.stderr.write('at %s value, "%s" is more than number of order in excel\n' % (number+1, data))
                    sys.exit(1)
            except ValueError as e:
                sys.stderr.write('at %s value, "%s" is not number\n' % (number+1, data))
                sys.exit(1)

    elif len(sys.argv) < 3:
        sys.stderr.write('Usage: %s\n' % (sys.argv[0]))
        sys.exit(1)

    if os.environ['DJANGO_SETTINGS_MODULE'] is None:
        os.environ['DJANGO_SETTINGS_MODULE'] = 'branded_express_projects.settings'

    _ecommerce = models.Ecommerce.objects.get(ecommerce_code=sys.argv[2])
    _merchant = models.Merchant.objects.get(merchant_id=sys.argv[3],ecommerce=_ecommerce)
    merchant=dict()
    merchant['merchant_id']=_merchant.merchant_id
    merchant['name']=_merchant.name
    merchant['address1']=_merchant.address1
    merchant['address2']=_merchant.address2
    merchant['sub_district']=_merchant.sub_district
    merchant['district']=_merchant.district
    merchant['city']=_merchant.city
    merchant['postal_code']=_merchant.postal_code
    merchant['country']=_merchant.country
    merchant['phone']=_merchant.phone
    merchant['email']=_merchant.email
    merchant['website']=_merchant.website
    #json.dumps(merchant)

    try:
        list_order=list()

        if sheet.ncols == ORDER_BATCH_FILE_HEADERS:
            # skip logo and header rows and start at first order row
            for row in range(sheet.nrows)[2:sheet.nrows]:
                order=dict()
                order['ecommerce_code']=_ecommerce.ecommerce_code

                #in case, value of column in excel is number
                if type(sheet.cell(row, 0).value)==float:
                    order['ecommerce_order_id']=int(sheet.cell(row, 0).value)
                else:
                    order['ecommerce_order_id']=sheet.cell(row, 0).value
                order['merchant']=merchant
                customer=dict()
                customer['customer_id']=sheet.cell(row, 1).value
                customer['name']=sheet.cell(row, 2).value
                customer['address1']=sheet.cell(row, 3).value
                customer['address2']=sheet.cell(row, 4).value
                customer['sub_district']=sheet.cell(row, 5).value
                customer['district']=sheet.cell(row, 6).value
                customer['city']=sheet.cell(row, 7).value
                customer['postal_code']=int(sheet.cell(row, 8).value)
                customer['country']=sheet.cell(row, 9).value
                customer['phone']=sheet.cell(row, 10).value
                customer['email']=sheet.cell(row, 11).value
                order['customer']=customer
                products=dict()

                #in case, value of column in excel is number
                if type(sheet.cell(row, 12).value)==float:
                    products['product_id']=int(sheet.cell(row, 12).value)
                else:
                    products['product_id']=sheet.cell(row, 12).value

                products['name']=sheet.cell(row, 13).value
                products['quantity']=1
                order['attribute']=sheet.cell(row, 15).value
                products['weight']=sheet.cell(row, 16).value
                products['width']=sheet.cell(row, 17).value
                products['height']=sheet.cell(row, 18).value
                products['depth']=sheet.cell(row, 19).value
                order['note']=sheet.cell(row,14).value
                order['products']=products
                order['shipping_method']="Standard"
                list_order.append(order)

            for number,order in enumerate(list_order):
                try:
                    if len(sys.argv) > 4:
                        # check parameter
                        for value in sys.argv[4:]:
                            # number+len(INCLUDE_HEADER) is started at third row in excel(first order)
                            if number+len(INCLUDE_HEADER)==int(value):
                                send_request(order)
                                break
                    else:
                        send_request(order)
                except urllib2.HTTPError, error:
                    log("Problem with order line %s \n"% (number+len(INCLUDE_HEADER)))
                    contents = error.read()
                    fout = open("error.html", "wb")
                    fout.write(contents)
                    fout.close

    except Exception as e:
        sys.stderr.write('Error :: %s \n' % (e.strerror))
        sys.exit(1)
