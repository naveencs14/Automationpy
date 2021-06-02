from O365 import Message,fluent_message, Attachment
from datetime import datetime, timedelta
import os
from orm.retrieve.hda_tvt_tracker import TvtTracker
import base64

import requests

from requests_ntlm import HttpNtlmAuth

class office_365():
    def __init__(self):
        self.content = ""

    def send_email(self, sender, password, receiver, subject, msg, cc_receciver=[]):

        authenticiation = (sender, password)
        m = fluent_message.Message(auth=authenticiation)
        m.setRecipients(receiver)
        m.setRecipients(cc_receciver, r_type='Cc')
        m.setSubject(subject)
        # att = Attachment(path=imgFile)
        # m.setBody(msg)
        m.setBodyHTML(msg)
        # m.attachments.append(att)
        # print m.json
        status = m.sendMessage('HDADesignAutomation@nbnco.com.au')
        # status = m.sendMessage()

        if status:
            pass
        else:
            print "Failed to send"

    def readHTML(self, path, htmlFile):
        file = open(os.path.join(path, htmlFile), 'r')
        content = file.read()
        file.close()
        return content

    def encodeImage(self):
        encoded = base64.b64encode(open(r'C:\TVT\New-Template\logo-3.png', "rb").read())
        return encoded

    def createMsg(self):
        a = TvtTracker()
        last_run_date = (datetime.now() - timedelta(days=0)).strftime('%d-%m-%Y')
        load_status = a.get_skynet_run_status(last_run_date)[0]['process_list'].split('|')
        sources = ''

        for x in load_status:
            f = x.split(':')
            sources += """  <tr>
                                            <td>{}</td>
                                            <td>{}</td>
                                            <td>{}</td>
                                        </tr>""".format(*f)

        head = """
                    <html>
                        <head>
                            <style>
                                table {
                                    font-family: arial, sans-serif;
                                    border-collapse: collapse;
                                    width: 30%;
                                }

                                td, th {
                                    border: 1px solid #dddddd;
                                    text-align: left;
                                    padding: 8px;
                                }

                                tr:nth-child(even) {
                                    background-color: #dddddd;
                                }
                            </style>
                        </head>

                """

        body = """<body>
                            <table>
                                <tr>
                                    <th>Source</th>
                                    <th>Sheet Name</th>
                                    <th>Status</th>
                                </tr>{}
                            </table>
                        </body>
                    </html>""".format(sources)

        msg = head + body
#         msg="""
#         <html>
# <head>
# <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
# </head>
# <body>
# <table class="table table-bordered">
#   <thead>
#     <tr>
#       <th scope="col">#</th>
#       <th scope="col">First</th>
#       <th scope="col">Last</th>
#       <th scope="col">Handle</th>
#     </tr>
#   </thead>
#   <tbody>
#     <tr>
#       <th scope="row">1</th>
#       <td>Mark</td>
#       <td>Otto</td>
#       <td>@mdo</td>
#     </tr>
#     <tr>
#       <th scope="row">2</th>
#       <td>Jacob</td>
#       <td>Thornton</td>
#       <td>@fat</td>
#     </tr>
#     <tr>
#       <th scope="row">3</th>
#       <td colspan="2">Larry the Bird</td>
#       <td>@twitter</td>
#     </tr>
#   </tbody>
# </table>
# </body>
# </html>
# """
        return msg


if __name__ == '__main__':
    c = office_365()
    a = TvtTracker()
    last_run_date = (datetime.now() - timedelta(days=0)).strftime('%d-%m-%Y')
    print a.get_skynet_run_status(last_run_date)[0]['process_list'].split('|')

    msg = c.createMsg()
