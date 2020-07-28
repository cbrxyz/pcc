from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import datetime

from helper import log

sheetUrl = r'SHEET URL GOES HERE'

def createHeadingRow():
    return ("<tr>" + 
                "<td style=\"background-color: white; text-align: center;\"></td>" +
                "<td style=\"background-color: lightgray; text-align: center;\">Files</td>" +
                "<td style=\"background-color: lightgray; text-align: center;\">Folders</td>" +
                "<td style=\"background-color: lightgray; text-align: center;\">Total</td>" +
            "</tr>")

def createCardRow(v1, v2, v3, v4):
    return ("<tr>" + 
                "<td style=\"background-color: lightgray; text-align: center;\">" + str(v1) + "</td>" +
                "<td style=\"background-color: #f1f1f1; text-align: center;\">" + str(v2) + "</td>" +
                "<td style=\"background-color: #f1f1f1; text-align: center;\">" + str(v3) + "</td>" +
                "<td style=\"background-color: #feffdd; text-align: center;\">" + str(v4) + "</td>" +
            "</tr>")
            
def createTotalRow(filesTotal, foldersTotal):
    return ("<tr>" + 
                "<td style=\"background-color: lightgray; text-align: center;\">Total</td>" +
                "<td style=\"background-color: #feffdd; text-align: center;\">" + str(filesTotal) + "</td>" +
                "<td style=\"background-color: #feffdd; text-align: center;\">" + str(foldersTotal) + "</td>" +
                "<td style=\"background-color: #FFE598; text-align: center;\">" + str(int(filesTotal) + int(foldersTotal)) + "</td>" +
            "</tr>")

def createOneColRow(v1, v2):
    return ("<tr>" +
                "<td colspan=\"3\" style=\"background-color: lightgray; text-align: center;\">" + str(v1) + "</td>" +
                "<td style=\"background-color: #f1f1f1;\">" + str(v2) + "</td>" 
            "</tr>")

def createComputerCard(info):
    res = ("<div class=\"computer-card\" style=\"height: auto; padding: 4px 0; text-align: center; width: 350px; border-radius: 5px ; margin: 3px; border: 1px solid black;\">" + 
                "<h2 style=\"margin: 5px 0;\">" + info['name'] + "</h2>" + 
                "<div style=\"text-transform: uppercase; font-size: 12px; background-color: #28a745; border-radius: 5px; color: white; display: inline-block; padding: 0 5px;\">Active</div>" + 
                "<h3 style=\"margin: 5px 0;\">Deleted Files/Folders</h3>" + 
                "<table style=\"width: 100%; margin-top: 5px;\">" + 
                    "<tbody>")
    res += createHeadingRow()
    res += createCardRow('Desktop', info['desktop-files-deleted'], info['desktop-folders-deleted'], str(int(info['desktop-files-deleted']) + int(info['desktop-folders-deleted'])))
    res += createCardRow('Documents', info['document-files-deleted'], info['document-folders-deleted'], str(int(info['document-files-deleted']) + int(info['document-folders-deleted'])))
    res += createCardRow('Pictures', info['pictures-files-deleted'], info['pictures-folders-deleted'], str(int(info['pictures-files-deleted']) + int(info['pictures-folders-deleted'])))
    res += createCardRow('Videos', info['videos-files-deleted'], info['videos-folders-deleted'], str(int(info['videos-files-deleted']) + int(info['videos-folders-deleted'])))
    res += createCardRow('Music', info['music-files-deleted'], info['music-folders-deleted'], str(int(info['music-files-deleted']) + int(info['music-folders-deleted'])))
    res += createCardRow('Downloads', info['downloads-files-deleted'], info['downloads-folders-deleted'], str(int(info['downloads-files-deleted']) + int(info['downloads-folders-deleted'])))
    res += createCardRow('Chrome', info['chrome-files-deleted'], info['chrome-folders-deleted'], str(int(info['chrome-files-deleted']) + int(info['chrome-folders-deleted'])))
    res += createCardRow('Edge', info['edge-files-deleted'], info['edge-folders-deleted'], str(int(info['edge-files-deleted']) + int(info['edge-folders-deleted'])))
    res += createTotalRow((
        int(info['desktop-files-deleted']) + 
        int(info['document-files-deleted']) +
        int(info['pictures-files-deleted']) +
        int(info['videos-files-deleted']) +
        int(info['music-files-deleted']) +
        int(info['downloads-files-deleted']) +
        int(info['chrome-files-deleted']) +
        int(info['edge-files-deleted'])
    ), (
        int(info['desktop-folders-deleted']) + 
        int(info['document-folders-deleted']) +
        int(info['pictures-folders-deleted']) +
        int(info['videos-folders-deleted']) +
        int(info['music-folders-deleted']) +
        int(info['downloads-folders-deleted']) +
        int(info['chrome-folders-deleted']) +
        int(info['edge-folders-deleted'])
    ))
    res += "<tr><td colspan=\"4\" style=\"height: 25px;\"></td></tr>"
    res += createOneColRow("Chrome instances killed", info['chrome-instances-killed'])
    res += createOneColRow("Edge instances killed", info['edge-instances-killed'])
    res += createOneColRow("Programs uninstalled", info['programs-uninstalled'])
    res += createOneColRow("Times ran", info['times-ran'])
    res += (        "</tbody>" + 
                "</table>" + 
            "</div>")
    return res

# information should be a dictionary of dictionaries containing computer information
# ex: [{'name': 'Computer 1', 'edge-instances-killed': 23, ...}, {'name': 'Computer 2', 'edge-instances-killed': 63, ...}, ...]
def sendMonthlyReport(api_key, from_e, to_e, subj, information):
    html = ("<div style=\"padding: 10px; font-size: 24px; border: 1px solid black; border-radius: 5px; text-align: center;\">" + 
            "Here is the monthly public computer cleaner report for the month of: " +
            "<b>" + (datetime.datetime.now() - datetime.timedelta(days=20)).strftime("%B") + "</b>" + 
            "</div>" + 
            "<div style=\"display: flex; justify-content: center;\">")
    for computer in information:
        html += createComputerCard(computer)
    html += (
        "</div>" + 
        "<a style=\"padding: 10px; font-size: 24px; border: 1px solid #28a745; border-radius: 5px; text-align: center; color: #28a745; display: block;\"" + 
        f"href=\"{sheetUrl}\">" +
        "Click here to view the managing Google Sheet." +
        "</a>")
    message = Mail(
        from_email=from_e,
        to_emails=to_e,
        subject=subj,
        html_content=html
    )
    try:
        sg = SendGridAPIClient(api_key)
        response = sg.send(message)
        if response.status_code == 202:
            log("Email was sent successfully.")
        else:
            log(f" [ERROR] Error sending SendGrid email. Email response was not 202. Instead: {response.status_code}")
    except:
        log(" [ERROR] Error sending SendGrid email. The SendGrid API client raised an exception.")