#https://www.youtube.com/watch?v=hyUw-koO2DA

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
        try:
            con = cx_Oracle.connect('system/elcaro')
            cur = con.cursor()

            # -----------Oracle Version ------------
            Oracle_Version = cur.execute("select * from v$version").fetchall()[0][0]
            sheet.update_acell('E' + str(rownum), Oracle_Version)

            # -----------Oracle Instance ------------
            text = ''
            text1 = ''
            text2 = ''
            text3 = ''
            Oracle_Instance = cur.execute("select sys_context('userenv','db_name') from dual").fetchall()[0][0]
            Oracle_PDBs = cur.execute("select PDB_NAME from DBA_PDBS where PDB_NAME !='PDB$SEED'").fetchall()
            Oracle_PDBs.insert(0, ('' + Oracle_Instance + '',))
            for Oracle_PDB in Oracle_PDBs:
                text += Oracle_PDB[0] + '\n\n'
                try:
                    con = cx_Oracle.connect('system/elcaro@' + Oracle_PDB[0])
                    cur = con.cursor()
                    result1 = cur.execute("SELECT LISTAGG(serviceday, ', ') WITHIN GROUP (ORDER BY serviceday) FROM (select distinct serviceday from bidb.sa_trips)").fetchall()[0][0]
                    result2 = cur.execute("SELECT LISTAGG(serviceday, ', ') WITHIN GROUP (ORDER BY serviceday) FROM (select distinct serviceday from bidb.sa_trips where sl_observed=1)").fetchall()[0][0]
                    text1 += "----" + Oracle_PDB[0] + "----\n\nScheduled Servicedays: " + result1 + "\n\nObserved Servicedays: " + result2 + "\n\n"
                    result3 = cur.execute("SELECT Customer, Branch, Patch_Date FROM BIDB.BI_Version").fetchall()[0]
                    text2 += "----" + Oracle_PDB[0] + "----\n\n" + result3[0] + "\n\n"
                    text3 += "----" + Oracle_PDB[0] + "----\n\nBranch - " + result3[1] + "\nPatch Date - " + result3[2] + "\n\n"
                except:
                    pass
            sheet.update_acell('F' + str(rownum), text)
            sheet.update_acell('G' + str(rownum), text1)
            sheet.update_acell('C' + str(rownum), text2)
            sheet.update_acell('D' + str(rownum), text3)
        except:
            pass
        # -----------Microstrategy Version ------------
        MSTR_Version = os.popen("mstrctl -s IntelligenceServer gs").read()
        MSTR_Version = ET.fromstring(MSTR_Version).find('./application/version').text
        text = MSTR_Version
        sheet.update_acell('H' + str(rownum), text) 

        # -----------Wildfly ------------
        text = '---Wildfly Folders---\n'
        for dir in os.listdir('D:/'):
            if dir.startswith('wildfly') or dir.startswith('Wildfly'):
                text += dir + '\n'

        path = str(os.popen("sc qc Wildfly | find \"BINARY_PATH_NAME\"").read())
        if path == "":
            text += "\nWildfly Service Unavailable\n\n"
        else:
            text += "\n---Wildfly Service Path---\n" + path[path.find(": ") + len(": "):path.rfind('')] + "\n\n"
            try:
                Wildfly = urllib.request.urlopen("http://localhost:8080").getcode()
            except:
                Wildfly = 0
            try:
                MTV = urllib.request.urlopen("http://localhost:8080/mtv").getcode()
            except:
                MTV = 0
            try:
                BIWEB = urllib.request.urlopen("http://localhost:8080/biweb").getcode()
            except:
                BIWEB = 0
            text += "Wildfly Running" if Wildfly == 200 else "Wildfly Unavailable\n"
            text += "MTV Running" if MTV == 200 else "MTV Unavailable\n"
            text += "BIWEB Running" if BIWEB == 200 else "BIWEB Unavailable"

        sheet.update_acell('I' + str(rownum), text)

        # -----------RAM ------------
        sheet.update_acell('J' + str(rownum), "Total: " + str(psutil.virtual_memory().total/1024000000)[:4] + " GB \n"
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
        sheet.update_acell('K' + str(rownum), text)

        # -----------Updated On ------------
        sheet.update_acell('L' + str(rownum), str(datetime.datetime.now())[:-7])

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

