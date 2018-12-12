# -*- coding: utf-8 -*-
"""
asvt_fingerprint.server

Very simple HTTP server in python.
Usage::
    ./dummy-web-server.py [<port>]
Send a GET request::
    curl http://localhost
        /?name=Dima --> add new user to the table
        /?id=1      --> enter/exit user with such id
        /?excel=    --> get excel table in browser

Send a HEAD request::
    curl -I http://localhost
Send a POST request::
    curl -d "name=Dima&id=1" http://localhost

Created by prudkovskiy on 29.11.18 23:50
"""
from datetime import datetime
import subprocess
from time import sleep
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse
import cgi
from excel_creator import create_new_employee, enter_employee, file_name
import xlrd
import csv

__author__ = 'prudkovskiy'


def csv_from_excel(filename):
    csv_file_name = filename.rstrip('.xlsx') + '.csv'
    wb = xlrd.open_workbook(file_name)
    sheet_name = datetime.now().strftime("%Y.%m")
    sh = wb.sheet_by_name(sheet_name)
    with open(csv_file_name, 'w') as f:
        wr = csv.writer(f, quoting=csv.QUOTE_ALL)

        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))
    return csv_file_name


class GP(BaseHTTPRequestHandler):
    def _set_headers(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()

    def do_HEAD(self):
        self._set_headers()

    def do_GET(self):
        self._set_headers()
        res = urlparse(self.path).query.replace('=', ' ').replace('&', ' ').split()
        if len(res) == 0:
            print('stop this shit')
            self.connection.close()
            return

        if res[0] == 'id':
            try:
                res[1] = int(res[1])

            except Exception as ex:
                res[1] = ord(res[1])
                # print(ex)

            print(res[1])
            enter_employee(file_name, res[1])
            self.wfile.write(b"Success!")
            sleep(2)
            self.connection.close()

        elif res[0] == 'name':
            name = res[1]
            print(name)
            id = create_new_employee(file_name, name)
            self.wfile.write(bytes(str(id).encode('utf-8')))
            sleep(3)
            self.connection.close()

        elif res[0] == 'excel':
            html_file = 'data.html'
            csv_file = csv_from_excel(file_name)
            bash_command = "excel2table {} {} -o".format(csv_file, html_file)
            subprocess.run('bash -c "source activate root; python -V"', shell=True)
            subprocess.call(bash_command.split())
            # p = subprocess.Popen(bash_command.split(), stdout=subprocess.PIPE,
            #                      stderr=subprocess.STDOUT, executable='/bin/bash')
            #
            # for line in p.stdout.readlines():
            #     print(line.decode('utf-8'))
            # p.wait()
            # sleep(1)
            with open(html_file, 'rb') as f:
                self.wfile.write(f.read())
            self.connection.close()

        else:
            self.connection.close()

    def do_POST(self):
        self._set_headers()
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={'REQUEST_METHOD': 'POST'}
        )
        print(form.getvalue("name"))
        print(form.getvalue("id"))
        self.wfile.write(b"<html><body><h1>POST Request Received!</h1></body></html>")


def run(server_class=HTTPServer, handler_class=GP, port=8080):
    ip_addr = '192.168.1.6'
    # ip_addr = '10.50.1.193'
    server_address = (ip_addr, port)
    httpd = server_class(server_address, handler_class)
    print(f'Server running at {ip_addr}:{port}...')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        httpd.server_close()


if __name__ == "__main__":
    print('Run server ({0})'.format(datetime.now()))
    run()
