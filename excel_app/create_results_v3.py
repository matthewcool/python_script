#!/usr/bin/python

import xmlrpclib
import json
import sys
import json
import xlrd
import xlutils.copy
import datetime
import os
from xlwt import Workbook, easyxf

inBook = xlrd.open_workbook('input_v3.xls', formatting_info=True)
outBook = xlutils.copy.copy(inBook)
dir_path = "/home/fujitsu/lava_report/{0}/".format(datetime.date.strftime(datetime.date.today(), '%Y%m%d'))

server = xmlrpclib.ServerProxy("http://matthew:7n0e6t63ymkxwx8omds2nuyzakp5deexew04pjnekleyr6drttqljdpnfsyde5aognuxnppucs1013q9kcx8pplnuxrwwzqk7qmmtc1phs00meonb30nr7iigejp8olg@192.168.1.75/RPC2/")

#column_config
A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R = range(18)

#id_config
id_config = {
    'kernelbuild' : 1,
    'device-tree' : 2,
    'ltp' : 3,
    'rt-ltp' : 4,
    'pwrmgmt' : 5,
    'aapits' : 6,
    'acpica' : 7,
    'acpi-abat' : 8,
    'fwts' : 9,
    'libhugetlbfs' : 10,
    'rt-hackbench' : 11,
    'perf' : 12,
    'gatortests' : 13,
    'e2eaudiotest' : 14,
    'smoke-tests-basic' : 15,
    'bootchart' : 16,
    'leb-basic-graphics' : 17,
    'sdcard' : 18,
    'usb-storage' : 19,
    'sata-storage' : 20,
    'emmc-storage' : 21,
}

#field_config
field_config = {
    'kernelbuild' : {'pass':{'column':E,'row':22},'fail':{'column':F,'row':22},'skip':{'column':G,'row':22},'unknown':{'column':H,'row':22}},
    'device-tree' : {'pass':{'column':E,'row':23},'fail':{'column':F,'row':23},'skip':{'column':G,'row':23},'unknown':{'column':H,'row':23}},
    'ltp' : {'pass':{'column':E,'row':24},'fail':{'column':F,'row':24},'skip':{'column':G,'row':24},'unknown':{'column':H,'row':24}},
    'rt-ltp' : {'pass':{'column':E,'row':25},'fail':{'column':F,'row':25},'skip':{'column':G,'row':25},'unknown':{'column':H,'row':25}},
    'pwrmgmt' : {'pass':{'column':E,'row':26},'fail':{'column':F,'row':26},'skip':{'column':G,'row':26},'unknown':{'column':H,'row':26}},
    'aapits' : {'pass':{'column':E,'row':27},'fail':{'column':F,'row':27},'skip':{'column':G,'row':27},'unknown':{'column':H,'row':27}},
    'perf' : {'pass':{'column':I,'row':33},'fail':{'column':J,'row':33},'skip':{'column':K,'row':33},'unknown':{'column':L,'row':33}},
    'gatortests' : {'pass':{'column':I,'row':34},'fail':{'column':J,'row':34},'skip':{'column':K,'row':34},'unknown':{'column':L,'row':34}},
    'e2eaudiotest' : {'pass':{'column':I,'row':35},'fail':{'column':J,'row':35},'skip':{'column':K,'row':35},'unknown':{'column':L,'row':35}},
    'smoke-tests-basic' : {'pass':{'column':I,'row':36},'fail':{'column':J,'row':36},'skip':{'column':K,'row':36},'unknown':{'column':L,'row':36}},
    'bootchart' : {'pass':{'column':I,'row':37},'fail':{'column':J,'row':37},'skip':{'column':K,'row':37},'unknown':{'column':L,'row':37}},
    'leb-basic-graphics' : {'pass':{'column':I,'row':38},'fail':{'column':J,'row':38},'skip':{'column':K,'row':38},'unknown':{'column':L,'row':38}},
    'sdcard' : {'pass':{'column':I,'row':39},'fail':{'column':J,'row':39},'skip':{'column':K,'row':39},'unknown':{'column':L,'row':39}},
    'usb-storage' : {'pass':{'column':I,'row':40},'fail':{'column':J,'row':40},'skip':{'column':K,'row':40},'unknown':{'column':L,'row':40}},
    'sata-storage' : {'pass':{'column':I,'row':41},'fail':{'column':J,'row':41},'skip':{'column':K,'row':41},'unknown':{'column':L,'row':41}},
    'emmc-storage' : {'pass':{'column':I,'row':42},'fail':{'column':J,'row':42},'skip':{'column':K,'row':42},'unknown':{'column':L,'row':42}},
}


def main():
    if len(sys.argv) == 1 or sys.argv[1] in {"-h", "--help"}:
        print("usage: {0} job_id [job_id [job_id [... job_id]]]".format(sys.argv[0]))
        sys.exit()

    sheet_index = 0
    for job_id in sys.argv[1:]:
        job_status = server.scheduler.job_status(job_id)
        if job_status['job_status'] == 'Complete':
            PASS, FAIL, SKIP, UNKNOWN, content_filename = get_bundle_results(job_status['bundle_sha1'])
            #print("{0} ->  pass:{1}, fail:{2}, skep:{3}, unknown:{4}".format(content_filename, PASS, FAIL, SKIP, UNKNOWN))
            if content_filename.startswith('ubuntu-desktop'):
                testcase = content_filename.split('(')[1].split(')')[0]
            else:
                testcase = content_filename
            outSheet = outBook.get_sheet(0)
            setOutCell(outSheet, field_config[testcase]['pass']['column'], field_config[testcase]['pass']['row'], PASS)
            setOutCell(outSheet, field_config[testcase]['fail']['column'], field_config[testcase]['fail']['row'], FAIL)
            setOutCell(outSheet, field_config[testcase]['skip']['column'], field_config[testcase]['skip']['row'], SKIP)
            setOutCell(outSheet, field_config[testcase]['unknown']['column'], field_config[testcase]['unknown']['row'], UNKNOWN)
            
            if FAIL != 0:
                sheet_index += 1
                add_fail_sheet(job_status['bundle_sha1'], sheet_index, testcase)
                get_logfiles(job_status['bundle_sha1'], testcase)

    outBook.save("output-{0}.xls".format(datetime.date.strftime(datetime.date.today(), '%Y%m%d')))    
           

def add_fail_sheet(sha1, sheet_index, testcase):
    outSheet = outBook.add_sheet("{0}. {1}".format(id_config[testcase], testcase))
    outSheet.col(2).width = outSheet.col(4).width = 6000
    outSheet.col(1).width = outSheet.col(3).width = 3000
    outSheet.write(1,1,'Fail Suite',easyxf(
    'font: name Arial, bold True;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour gray25;'
    'alignment: horizontal center, vertical center, wrap True;'))
    outSheet.write(1,2,testcase,easyxf(
    'font: name Arial;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'alignment: horizontal center, vertical center, wrap True;'))
    outSheet.write(3,1,'No',easyxf(
    'font: name Arial, bold True;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour gray25;'
    'alignment: horizontal center, vertical center, wrap True;'))
    outSheet.write(3,2,'Test case',easyxf(
    'font: name Arial, bold True;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour gray25;'
    'alignment: horizontal center, vertical center, wrap True;'))
    outSheet.write(3,3,'Fails',easyxf(
    'font: name Arial, bold True;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour gray25;'
    'alignment: horizontal center, vertical center, wrap True;'))
    outSheet.write(3,4,'Measurement',easyxf(
    'font: name Arial, bold True;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour gray25;'
    'alignment: horizontal center, vertical center, wrap True;'))
    
    bundle_content = server.dashboard.get(sha1)
    content = json.loads(bundle_content['content'])
    case_amount = len(content['test_runs'][0]['test_results'])
    row = 4
    for i in range(0, case_amount):
        test_case_id = content['test_runs'][0]['test_results'][i]['test_case_id']
        result = content['test_runs'][0]['test_results'][i]['result']
        if result == 'fail':
            outSheet.write(row,B,i+1,easyxf(
            'font: name Arial;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'alignment: horizontal center, vertical center, wrap True;'))
            outSheet.write(row,C,test_case_id,easyxf(
            'font: name Arial;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'alignment: horizontal center, vertical center, wrap True;'))
            outSheet.write(row,D,'fail',easyxf(
            'font: name Arial;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'alignment: horizontal center, vertical center, wrap True;'))
            outSheet.write(row,E,'Not specified',easyxf(
            'font: name Arial;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'alignment: horizontal center, vertical center, wrap True;'))
            row += 1  
 

def get_logfiles(sha1, testcase):
    case_dir_path = "{0}{1}/".format(dir_path, testcase)
    if not os.path.exists(dir_path): os.makedirs(dir_path)
    if not os.path.exists(case_dir_path): os.makedirs(case_dir_path)
    
    bundle_content = server.dashboard.get(sha1)
    content = json.loads(bundle_content['content'])
    log_amount = len(content['test_runs'][0]['attachments'])
    for i in range(0, log_amount):
        pathname = content['test_runs'][0]['attachments'][i]['pathname']
        base64_content = content['test_runs'][0]['attachments'][i]['content']
        data = base64_content.decode('base64', 'strict')
        f = open(case_dir_path+pathname, 'w')
        try:
            f.writelines(data)
        finally:
            f.close()


def get_bundle_results(bundle_sha1):
    PASS = FAIL = SKIP = UNKNOWN = 0
    bundle_content = server.dashboard.get(bundle_sha1)
    content_filename = bundle_content['content_filename']
    content = json.loads(bundle_content['content'])
    case_amount = len(content['test_runs'][0]['test_results'])
    for i in range(0, case_amount):
        test_case_id = content['test_runs'][0]['test_results'][i]['test_case_id']
        result = content['test_runs'][0]['test_results'][i]['result']
        if result == 'pass':
            PASS += 1
        elif result == 'fail':
            FAIL += 1
        elif result == 'skip':
            SKIP += 1
        else:
            UNKNOWN += 1
    return PASS, FAIL, SKIP, UNKNOWN, content_filename


def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell


def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
    # END HACK

main()

