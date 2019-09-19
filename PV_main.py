from tkinter import *
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import colorchooser
from tkinter import filedialog
from xml.dom import minidom
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from requests_ntlm import HttpNtlmAuth
from cryptography.fernet import Fernet
import time, os, sys, json, requests, shutil, re, smtplib, urllib, codecs

# =========== Version ?.?.? ================
sVersion = '1.0.7'

# ======== Parse an xml file by name =========
mydoc = minidom.parse('configuration\config.xml')
xServer = mydoc.getElementsByTagName('server')
xMail = mydoc.getElementsByTagName('mail')
xbitbucket = mydoc.getElementsByTagName('key')
# ====== Get parametric from XML file =======
xSever = xServer[0].firstChild.data
xUser = xServer[1].firstChild.data
xPassword = xServer[2].firstChild.data
xServerUp = xServer[3].firstChild.data
xSMTP = xServer[4].firstChild.data
xMailTo = xMail[0].firstChild.data
xMailCc = xMail[1].firstChild.data
xMailSub = xMail[2].firstChild.data
xSender = xMail[3].firstChild.data
xReceivers = xMail[4].firstChild.data

kLink = xbitbucket[0].firstChild.data
kUser = xbitbucket[1].firstChild.data
kRepos = xbitbucket[2].firstChild.data
kTagsLimit = xbitbucket[3].firstChild.data


# ==========================================

def pClose():
    status = messagebox.askyesno(title="Close", message="Close Program?")
    if status > 0:
        sys.exit()


# ============= Get Tags From Bitbucket ==================
def gTags():
    # ===================== Read Json file ============================
    readeJson()

    global kPass0, kUser0
    # ============== Get Test code package from Bitbucket ==========
    root.update_idletasks()
    kUser1 = simpledialog.askstring(title=" Bitbucket ", prompt="UserName:", parent=root, initialvalue=kUser)
    if kUser1 is None: exit();

    kPass = simpledialog.askstring(title=" Bitbucket ", prompt="Password:", show='*')
    if kPass is None: exit();

    kPass0 = kPass
    kUser0 = kUser1
    headers = {'Content-type': 'application/json'}
    url = kLink+'/repos/script/tags?limit=' + kTagsLimit + ''

    r2 = requests.get(url, headers=headers, auth=(kUser, kPass), timeout=50)
    jdata0 = r2.json()
    ddBit = []
    for i in range(jdata0["size"]):
        ddBit.append(jdata0["values"][i]["displayId"]);

    return ddBit
    # ==================================================================


def fCon(selected_item):
    global sInputPKG
    sInputPKG = str(selected_item.strip())


def fOk(event):
    # ===================== Read Json file ============================
    readeJson()

    refT = 'refs%2Ftags%2F'
    headers = {'Content-type': 'application/octet-stream'}
    url = kLink + '/repos/' + kRepos + '/archive?at=' + refT + '' + sInputPKG + '&format=zip'
    print(url)

    r2 = requests.get(url, headers=headers, auth=(kUser, kPass0), timeout=50)
    temp = 'temp/' + sInputPKG + '.zip'

    if r2.status_code == 200:
        open(temp, 'wb').write(r2.content)
        baseName = temp.rsplit("/", 1)
        beName = baseName[1].split("_")
        afName = beName[4].split(".")
        reName = 'temp/QR__' + beName[2] + '' + afName[0] + '' + afName[1] + '' \
                 + afName[2] + '' + afName[3] + '_' + beName[3] + '_' + beName[2] + '.zip'
        if os.path.isfile(reName) is True: os.unlink(reName)
        os.rename(temp, reName)
        print("Success: Download from Bitbucker.")
    elif r2.status_code == 401:
        print("Login Fail: Unable to login. Please check username in config.xml")
        l_res1.insert(1, "Warning: Login Fail. Please check username in config.xml")
        eLogin = messagebox.showwarning("Warning", "Something Wrong! Please try again.")
        if eLogin == "ok": sys.exit()
    else:
        print("Error: Can not download from Bitbucker.")
        l_res1.insert(1, "Error: Can not download from Bitbucker.")
        eDown = messagebox.showwarning("Warning", "Something Wrong! Please try again.")
        if eDown == "ok": sys.exit()

    # =================== Get last file name =========================
    baseName0 = reName.rsplit("/", 1)  # fullPath + nameFile
    bName = baseName0[1].split(".")  # remove .zip
    print(bName[0])

    bName2 = bName[0].split("_")  # Example:
    bNameLen = len(bName2)
    bType = bName2[bNameLen - 1]  #
    bProduct = bName2[bNameLen - 2]  #

    # ================== Update Json file (write file) ===============
    jdata["package"]["requester_email"] = xMailTo
    jdata["package"]["master_revision"] = bName[0]
    jdata["package"]["file_path"] = xServerUp
    jdata["package"]["base_code"] = bName[0][0:15]
    jdata["package"]["crm_id"] = bName[0][0:15]
    writeJson()
    # =================================================================

    # ================= Copy file to another directory ================
    try:
        copyTo = shutil.copy(reName.replace("/", "//"), jdata["package"]["file_path"].replace("\"", "\"\""))
        print("Success: Copied file to server. : ", copyTo)
    except Exception:
        print("Error: Unable Copy file to Server.")
        exit()

    # ==================================================================

    # ================= Request Http send Json =========================
    header_content = {'Content-type': 'application/json'}
    readeJson()
    ddata = json.dumps(jdata)
    r = requests.post(xSever, auth=(xUser, xPassword), headers=header_content, data=ddata)

    if r.status_code == 200:
        rStatus = "<font size=\"3\" color=\"green\">Success</font>"
    else:
        rStatus = "<font size=\"3\" color=\"red\">Fail</font>"
    # ==================================================================

    # =================== Set result file path ==========================
    resutlFilePath = jdata["package"]["file_path"] + "\\PKGS\\" \
                     + bType + "\\" + bProduct + "\\" + bName[0][0:15] + "\\" \
                     + bName[0][0:15] + "_" + bProduct + "_" + bType + "_result.html"
    print(resutlFilePath)

    while True:
        print("Waiting: PV background process.")
        cFile = os.path.isfile(resutlFilePath)
        if cFile == True:
            print("Success: PV process is completed.")
            f = codecs.open(resutlFilePath, 'r')
            if 'No rule' in f.read():
                rStatusRule = "<font size=\"3\" color=\"green\">Success</font>"
            else:
                rStatusRule = "<font size=\"3\" color=\"red\">Fail</font>"

            try:
                if os.path.isfile(jdata["package"]["file_path"] + "\\" + baseName0[1]) is True:
                    os.unlink(jdata["package"]["file_path"] + "\\" + baseName0[1])
                if os.path.isfile(reName) is True:
                    os.unlink(reName)
                print("Success: Delete file on local and server.:" + baseName0[1])
            except Exception:
                print("Error: Unable to Move file." + baseName0[1])
            break;
        else:
            time.sleep(10)
    # ======================== Result status(1)===========================
    l_res1.insert(1, "Revision: " + jdata["package"]["master_revision"])
    l_res1.insert(2, "Requester Email: " + jdata["package"]["requester_email"])
    l_res1.insert(3, "Email cc: " + xMailCc)
    l_res1.insert(4, "File_path: " + jdata["package"]["file_path"])
    l_res1.insert(5, "\n  ")
    l_res1.insert(6, "Result.... ")
    l_res1.insert(7, ">> " + str(r.content))
    ##========================= Send mail ==============================
    sender = (xSender)
    receivers = [xReceivers]

    msg = MIMEMultipart('alternative')
    msg['Subject'] = xMailSub
    msg['From'] = "PV Upload and Check <PV_System@wdc.com>"
    msg['To'] = xMailTo
    msg['Cc'] = xMailCc

    html = """\
<html>
<head></head>
<body>
<strong>Hello PV User,</strong> <br />
<p style="text-indent: 2.5em;">""" + rStatus + """: PV Upload Status</p>	
<p style="text-indent: 2.5em;"><font size=\"1\">Details    : """ + str(r.content) + """</font></p>	
<p style="text-indent: 2.5em;"><font size=\"1\">Revision   : """ + jdata["package"]["master_revision"] + """</font> </p>	
<p style="text-indent: 2.5em;">-----------------------------------------------</p>
<p style="text-indent: 2.5em;">""" + rStatusRule + """: Package Validation Rule</p>
<p style="text-indent: 2.5em;"><font size=\"1\">File path  : """ + resutlFilePath + """</font></p>	


<br /><br />	   
<strong>Thank you.</strong><br />
<strong>PV Admin</strong> <br />
<br />
<strong>******* No reply *********</strong>	<br /> 

</body>

</html>"""

    part = MIMEText(html, 'html')
    msg.attach(part)

    try:
        smtpObj = smtplib.SMTP(xSMTP)
        smtpObj.sendmail(sender, receivers, msg.as_string())
        smtpObj.quit()
        b_ok["state"] = "normal"
        print("Success: Send Email.")

        status0 = messagebox.askyesno(title="Success", message="Success: Close Program?")
        if status0 > 0: sys.exit()
    except smtplib.SMTPException:
        print("Error: Unable to send email.")


# ==================================================================

# ========= Select test code package file ===============
def mAbout():
    status = messagebox.showinfo("About me.",
                                 "Build time: Sep 19 2019 - 11:40:50 \n\nPV Package Validator."
                                 + sVersion + " @2019\nDevelop:      Kijjasarun K. \nCompany:    WDC Thailand.")


# ================ Read Json file =======================
def readeJson():
    with open('configuration\post.json', 'r') as f:
        global jdata
        jdata = json.load(f)


# ================ Write Json file =======================
def writeJson():
    with open('configuration\post.json', 'w') as outfile: json.dump(jdata, outfile)


# ==================== Start Main =======================
root = Tk()
root.geometry('650x400')
root.title("PV Upload/Check :Version " + sVersion)

# ==================== File Menu ========================
menubar = Menu(root)
fileMenu = Menu(menubar, tearoff=0)
fileMenu.add_command(label="Close", command=pClose)

# ====================== Help Menu =======================
helpMenu = Menu(menubar, tearoff=0)
helpMenu.add_command(label="About", command=mAbout)

menubar.add_cascade(label="File", menu=fileMenu)
menubar.add_cascade(label="Help", menu=helpMenu)

root.config(menu=menubar)
label = Label(root)
label.grid()

# ====================== GUI Layout ======================
l_hello = Label(text="PV Package Validator " + sVersion + " ", fg="blue", font="Verdana 10 bold")
l_line2 = Label(text="---")
l_tcp = Label(text="Test code Package: ", font="Verdana 8 bold")
str1 = StringVar()
padded_choices = gTags()
max_len = 38
tagVar = StringVar()
e_tcp = OptionMenu(root, tagVar, padded_choices[0], *padded_choices, command=fCon)
e_tcp.config(width=45)
# ++++++++++++++++++++++++++++++++++++++++

b_ok = Button(text="OK", width=8)
b_ok.bind('<Button-1>', fOk)
b_cal = Button(text="Cancel", width=8, command=pClose)
l_line3 = Label(text="Result status...")
l_res1 = Listbox(width=50)
l_line4 = Label(text="\"PV Package Validator." + sVersion + " @2019\"", font="Verdana 6")

# ======================== Grid Layout ========================
l_hello.grid(row=0, columnspan=4)
l_line2.grid(row=1, columnspan=4)
l_tcp.grid(row=2)
e_tcp.grid(row=2, column=1, columnspan=2, padx=3, pady=3)

l_line3.grid(row=3, column=1, sticky=W)
l_res1.grid(row=4, column=1, columnspan=2, padx=3, pady=3)

# ================= OK and Cancel (row ++1)=====================
b_ok.grid(row=5, column=1, sticky=E, padx=3, pady=3)
b_cal.grid(row=5, column=2, sticky=W, padx=3, pady=3)

l_line4.grid(columnspan=4, sticky=E)

root.mainloop()
# ===============================================================
