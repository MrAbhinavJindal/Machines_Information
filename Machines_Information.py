#https://www.youtube.com/watch?v=hyUw-koO2DA

import gspread, datetime, socket, string, psutil, os, cx_Oracle, subprocess
from oauth2client.service_account import ServiceAccountCredentials
from ctypes import windll

client_secret = {
  "type": "service_account",
  "project_id": "datauploader-357004",
  "private_key_id": "9a1a8037eb0b64578064cc4adcb99c75b67ad093",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCjRPrRvLN904Vp\nj+eYAcabNzg0YkxN94auF2BQKWsUpynEcFGuzT8RhaWKMBgqbuycrrDXWTtQssTg\nIMzrU1EBoaZTFo+ggYi+/AA9wfKkf37oQqiOHNN9a0ZPwmsX2ZCXhC6wOafzvwHs\n/yvyducge6wDBFs6LPrLG66ETHqFIdCO1KXTcim8UhYFTi1BIQtExpNM0b2EyLZ+\nFf6D9TZl30ENTj/eU5Y+MCxlui8H39RM6tNLOqvo6zt6XKpRiVFCOm9mLGWUrazW\nK3DSbLauNntTKzpAMAT6paGfHahVFzn7I3Swtmnd36lpq7TZv0Uv3eYHQbijZc92\nKDxKFkdXAgMBAAECggEAGzHpW1fEnU/xuBhrBDPrhRK+Dw+pKRAdQ/s1z4m5QdU+\n4kRZKWDGQWC9puuwOBqGLpT0dMedVymScciqNAUKiFczcUHf8OVn8HPg2xM78NrM\nZIXR4JX8Q4xDnRPDqyh8lULu/0B4nA4z6moYbl26zXt8C9FFHBwTjBLozzWT73+4\nVu3bD1DQcyOrjHO861f1vVcoh+jKX4+tZfnEZkP0YnFhEBGNWDPyPtl0yd9ThPNx\ndskGmi26YVwfEgXFnvAl+dJlSKkaoueewIuW4kskQjHpCRVJVXvbC85qlI7nNWmn\noHp5P04LeGVg4LuYww2vuzF/W/KW5MrFoDhEBG0fpQKBgQDZDvmu+EPnbMuOWYq6\n2yYAwaVXuxKaltnLe6UjQsH7xRHt4Da3BcrrQGSTTIDAk5eaJsAO4uArU/Ru3nL1\nhXfeEHKelRGKB9Jf6obXmyIclBEjckpB7j7KmcpkY1fSs57iIu6FN6rkywctx3h2\nd/W3nJ54geLl23y6kLZOSHc+0wKBgQDAj5cwmPLsrcIqoNsQxJ4u3d2rDFLhveRL\noj41SbhnBC3h2iQ1TEaOOqOoL40k22gphkN32HaZgjvvj1+96zKTweuLoDoVhqLQ\nX6059HtFNDFMUSgkqRNdZjhIFpREenY4SJl0bqEuOPgy0nC+yHKAmv48RjKB+0/r\nDh7G6oOq7QKBgF9OXAOfrvEmrBpM5sU1BHLAlED5Oyn1opveJpxc66AI39563Itw\nV7EEDSVAKihkpeRhr2LZ62Qa8PDda8yyVfeDcVCAU7svxAepipuQ2mGCAiR2QnTA\nj4GWFXAOzrkNdW4FuIV18+uR2g0X0KTz90gv1MVFAsO6pAGnGOU2nGVRAoGAUU8Q\nfzPGN8wzFb7wYYc0aAPFKwm8IZgGQy2R6PxlAhLQsPJkoaDAliQKoOTbS3nd5NLN\nwFhF1BIa7s/ylIYwyBV1OXMBs78zFpuf0L38Iz+jpV8LfVdrVt/n2gC2wKeZLbDy\nIyjnpFXn78XOV7DaMJXBzn+xqhMNLuq6cjHqQQ0CgYAhRsN/6KOBNYjYNHCRndf4\nOVoX8MOkUrYomXjb/lxEA3HuhbE8wnwrOhFWBrg1foB+u6EeXZyUr1uLpnTl3CCJ\nWPWkThn+63Qberx9igoI3ukqjKm2pSpQK3NxIcuxnx+hoTyd6KLgPvqnmgG9ng30\nIHccnKsKlvQcJSI9uPNFFA==\n-----END PRIVATE KEY-----\n",
  "client_email": "machinesinfo@datauploader-357004.iam.gserviceaccount.com",
  "client_id": "110103315677119763765",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/machinesinfo%40datauploader-357004.iam.gserviceaccount.com"
}

print(socket.gethostname().upper())
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(client_secret, scopes=scope)

file = gspread.authorize(creds)
workbook = file.open("Machines Information")
sheet = workbook.sheet1

rownum = 1
for cell in sheet.range('B2:B20'):
    rownum += 1
    if cell.value.upper() == socket.gethostname().upper():

        sheet.update_acell('C' + str(rownum), "")
        sheet.update_acell('D' + str(rownum), "")
        sheet.update_acell('E' + str(rownum), "")
        sheet.update_acell('F' + str(rownum), "")
        sheet.update_acell('G' + str(rownum), "")
        sheet.update_acell('H' + str(rownum), "")
        sheet.update_acell('I' + str(rownum), "")
        sheet.update_acell('J' + str(rownum), "")
        sheet.update_acell('K' + str(rownum), "")
        sheet.update_acell('L' + str(rownum), "")

        try:
            # -----------Oracle Version ------------
            con = cx_Oracle.connect('system/elcaro')
            cur = con.cursor()
            Oracle_Version = cur.execute("select * from v$version").fetchall()[0][0]
            sheet.update_acell('E' + str(rownum), Oracle_Version)

            # -----------Oracle Instance ------------
            text = ''
            text1 = ''
            text2 = ''
            text3 = ''

            domain_name = cur.execute("select case when display_value is null then '' else display_value end from v$parameter where name ='db_domain'").fetchall()[0][0]
            domain_name = "" if domain_name is None else "." + domain_name
            print("domain_name - " + domain_name)
            Oracle_CDB = cur.execute("select sys_context('userenv','db_name') from dual").fetchall()
            print("Oracle_CDB -" + str(Oracle_CDB))
            Oracle_PDBs = cur.execute("select PDB_NAME from DBA_PDBS where PDB_NAME !='PDB$SEED'").fetchall()
            print("Oracle_PDBs - " + str(Oracle_PDBs))

            if Oracle_PDBs:
                Instances = Oracle_PDBs
            else:
                Instances = Oracle_CDB

            for Instance in Instances:
                print(Instance[0] + domain_name)
                text += Instance[0] + domain_name + '\n\n'
                con = cx_Oracle.connect('system/elcaro@' + Instance[0] + domain_name)
                cur = con.cursor()
                result1 = cur.execute("SELECT LISTAGG(serviceday, ', ') WITHIN GROUP (ORDER BY serviceday) FROM (select distinct serviceday from bidb.sa_trips)").fetchall()[0][0]
                result2 = cur.execute("SELECT LISTAGG(serviceday, ', ') WITHIN GROUP (ORDER BY serviceday) FROM (select distinct serviceday from bidb.sa_trips where sl_observed=1)").fetchall()[0][0]
                text1 += "----" + Instance[0] + domain_name + "----\n\nScheduled Servicedays: " + result1 + "\n\nObserved Servicedays: " + result2 + "\n\n"
                result3 = cur.execute("SELECT Customer, Branch, Patch_Date FROM BIDB.BI_Version").fetchall()[0]
                text2 += "----" + Instance[0] + domain_name + "----\n\n" + result3[0] + "\n\n"
                text3 += "----" + Instance[0] + domain_name + "----\n\nBranch - " + result3[1] + "\nPatch Date - " + result3[2] + "\n\n"
            sheet.update_acell('F' + str(rownum), text.rstrip('\n\n'))
            sheet.update_acell('G' + str(rownum), text1.rstrip('\n\n'))
            sheet.update_acell('C' + str(rownum), text2.rstrip('\n\n'))
            sheet.update_acell('D' + str(rownum), text3.rstrip('\n\n'))
        except cx_Oracle.DatabaseError as e:
            sheet.update_acell('C' + str(rownum), "Oracle Database Not Installed")
            sheet.update_acell('D' + str(rownum), "Oracle Database Not Installed")
            sheet.update_acell('E' + str(rownum), e)
            sheet.update_acell('F' + str(rownum), "Oracle Database Not Installed")
            sheet.update_acell('G' + str(rownum), "Oracle Database Not Installed")


        # -----------Microstrategy Version ------------
        p = subprocess.Popen("mstrctl -s IntelligenceServer gs | find \"<version>\"", stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        output, error = p.communicate()
        if output.decode() == "":
            text = "Microstrategy Not Installed"
        else:
            MSTR_Version = output.decode()
            text = MSTR_Version[MSTR_Version.find("<version>") + len("<version>"):MSTR_Version.rfind("</version>")]
        sheet.update_acell('H' + str(rownum), text)

        # -----------JAVA Version ------------
        JAVA_Version = os.popen('java -version 2>&1 | find \"version\"').read()
        text = JAVA_Version
        if text == "":
            text = "JAVA Not Installed"
        sheet.update_acell('I' + str(rownum), text)

        # -----------Wildfly ------------
        text = '---Wildfly Folders---\n'
        for dir in os.listdir('D:/'):
            if dir.startswith('wildfly') or dir.startswith('Wildfly'):
                text += dir + '\n'
                os.chdir('D:/' + dir + '/bin')
                p = subprocess.run("jboss-cli.bat -c deployment-info", stdout=subprocess.PIPE, stdin=subprocess.PIPE)
                Result1 = p.stdout.splitlines()[0].decode()
                if Result1.startswith("Failed") or Result1.startswith("Press"):
                    text += 'Wildfly Not running\n\n'
                else:
                    Result2 = p.stdout.splitlines()[1].decode().split()
                    text += Result2[0] + " - " + Result2[4] + "\n\n"

        path = str(os.popen("sc qc Wildfly | find \"BINARY_PATH_NAME\"").read())
        if path == "":
            text += "Wildfly Service Unavailable\n\n"
        else:
            text += "---Wildfly Service Path---\n" + path[path.find(": ") + len(": "):path.rfind('')] + "\n\n"

        sheet.update_acell('J' + str(rownum), text)

        # -----------RAM ------------
        sheet.update_acell('K' + str(rownum), "Total: " + str(psutil.virtual_memory().total/1024000000)[:4] + " GB \n"
                                              "Used: " + str(psutil.virtual_memory()[3]/1024000000)[:4] + " GB")

        # -----------Drive Info ------------
        text = ''
        bitmask = windll.kernel32.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                try:
                    hdd = psutil.disk_usage(letter + ":")
                    text += "---" + letter + " Drive---\n" \
                           "Total: " + str(int(hdd.total / (2 ** 30))) + " GB\n" \
                           "Free: " + str(int(hdd.free / (2 ** 30))) + " GB\n\n"
                except:
                    pass
            bitmask >>= 1
        sheet.update_acell('L' + str(rownum), text.rstrip("\n\n"))

        # -----------Updated On ------------
        sheet.update_acell('M' + str(rownum), str(datetime.datetime.now())[:-7])

"""
Oracle_Version = subprocess.Popen(["sqlplus", "//"], stdout=subprocess.PIPE, shell=True).communicate()
Oracle_Version = (str(Oracle_Version).split('Version '))[1].split(r'\r\n\r\nCopyright')[0]
print("Oracle Version = " + Oracle_Version)

for child in MSTR_Version:
    print(child.tag, " = ", child.text)

# "Used: " + str(int(hdd.used / (2 ** 30))) + " GB\n" \

# print(sheet.acell('A2').value)
# print(sheet.cell(1, 1).value)
# print(sheet.row_values(2))
# sheet.update_acell('A2','zxczxc')
"""
