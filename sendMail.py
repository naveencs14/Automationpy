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

    # img = c.encodeImage()
    # print img
    # c.send_email('ajitabhsidhan@nbnco.com.au', 'Mrvi$$44yo', ['ajitabhsidhan@nbnco.com.au', 'ancyatomy@nbnco.com.au', 'adrianaksan@nbnco.com.au'], 'Test', 'Test Message')
    # c.send_email('ajitabhsidhan@nbnco.com.au', 'Mrvi$$44yo',
    #              [
    #                  'ajitabhsidhan@nbnco.com.au',
    #                  #  'HDADesignAutomation@nbnco.com.au',
    #                  # 'anujalal@nbnco.com.au',
    #                  # 'carinechane@nbnco.com.au',
    #                  # 'deenatom@nbnco.com.au',
    #                  # 'indikailukkumbure@nbnco.com.au',
    #                  # 'jeffchan@nbnco.com.au',
    #                  # 'jocasserly@nbnco.com.au',
    #                  # 'nathanenglish2@nbnco.com.au',
    #                  # 'patrickcoffey@nbnco.com.au',
    #                  # 'rajivramani@nbnco.com.au',
    #                  # 'samariakia@nbnco.com.au',
    #                  # 'shaunamorris@nbnco.com.au',
    #                  # 'shraddhachaubal@nbnco.com.au',
    #                  # 'taraporter@nbnco.com.au',
    #                  # 'weijunzhang@nbnco.com.au',
    #                  # 'mattwilks@nbnco.com.au',
    #                  # 'lancekeech@nbnco.com.au',
    #                  # 'ahmadserhan@nbnco.com.au',
    #                  # 'jeremyjmastop@nbnco.com.au',
    #                  # 'arujunangunasingam@nbnco.com.au',
    #                  # 'salilsalahuddin@nbnco.com.au',
    #                  # 'rupeshmuraleedharan@nbnco.com.au',
    #                  # 'angelicaleong@nbnco.com.au',
    #                  # 'jaredbalstrup@nbnco.com.au',
    #                  # 'jasonjudd1@nbnco.com.au',
    #                  # 'michaelcvetkovic@nbnco.com.au',
    #                  # 'chittetteraj@nbnco.com.au',
    #                  # 'teaguepetersen@nbnco.com.au',
    #                  # 'grahammcevoy@nbnco.com.au',
    #              ],
    #              # '[TVT Release v3.23]',
    #              'SkyNet DataSources Load - {}'.format(datetime.now().strftime('%d-%m-%Y')),
    #              msg,
    #              [
    #                  # 'shilpajhunjhunwala@nbnco.com.au',
    #                  #  'BradBainger@nbnco.com.au',
    #                  #  'georgebekes@nbnco.com.au',
    #                  #  'krishnakiran@nbnco.com.au',
    #                  #  'HDA_Downstream_Completion@nbnco.com.au',
    #                  #  'HDA_Design_Automation@nbnco.com.au',
    #                  #  'balajiranganathan@nbnco.com.au',
    #              ]
    #              )

    # current_session = requests.Session()
    # # turn off the SSL warning
    # current_session.verify = False
    # # validate the credentials
    # current_session.auth = HttpNtlmAuth('ajitabhsidhan', 'Mrvi$$44yo', current_session)
    #
    # aaa = current_session.get('https://outlook.office365.com/api/v1.0/me/messages')
    #
    # print aaa
