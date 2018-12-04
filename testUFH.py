import pyodbc, os, re, shutil, glob
from ftplib import FTP
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import subprocess

cnxn = pyodbc.connect('Driver={SQL Server};'
                      'Server=SQL2;'
                      'Database=ClaimsGeneral;'
                      'Trusted_Connection=yes;')


def pgp(io_status_param, DestinationFilemask, PGPKeyID, PGPKeyPassword, DestinationFileExtension):
    if io_status_param == 'I':
        # decrypt --call cmd line through subprocess method and paste cmd text
        InputFilePath = ''  # ((String)globalMap.get("tFileList_1_CURRENT_FILEPATH"))

        cmd = "cmd /c c:\\Talend\\GnuPG2\\bin\\gpg.exe --pinentry-mode=loopback --passphrase \""\
              + PGPKeyPassword + "\" -d -o \"" + InputFilePath[0:len(InputFilePath)-4]\
              +"\" " + InputFilePath
        subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).wait()
        pass

    elif io_status_param == 'O':
        # encrypt --call cmd line through subprocess method and paste cmd text
        FileDirectory = ''            # ((String)globalMap.get("tFileList_1_CURRENT_FILEDIRECTORY"))
        OutputFilename = '' + ".pgp"  # ((String)globalMap.get("tFileList_1_CURRENT_FILE")) + ".pgp"
        KeyID = PGPKeyID
        InputFilePath = ''            # ((String)globalMap.get("tFileList_1_CURRENT_FILEPATH"))

        cmd = "cmd /c c:\\Talend\\BU_GNUPG\\gpg.exe -o " + FileDirectory + "/" + OutputFilename + " -r \"" \
              + KeyID + "\" -e " + InputFilePath
        subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE).wait()
        pass


def modify(SourceFilemask, SourceFilemask2, SourceFilemask3, DestinationFilemask,
                   DestinationFileExtension, LocalPath, ConcatenateMultiFiles, ConcatenateHasHeader,
                   FileFormatIndicator, DateStampFile, entity):
    # rename files with date time stamp, and numeric.sequence()
    if (ConcatenateMultiFiles == 0) and (ConcatenateHasHeader == 0) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi")):
        pass

    # Rename Concatenated TEXT/CSV Files INCLUDE Header
    if (ConcatenateMultiFiles == 1) and (ConcatenateHasHeader == 1) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi")):
        pass

    # Rename Concatenated TEXT/CSV Files EXCLUDE Header
    if (ConcatenateMultiFiles == 1) and (ConcatenateHasHeader == 0) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi")):
        pass

    # Rename Concatenated EXCEL Files
    if (ConcatenateMultiFiles == 1) and (FileFormatIndicator.equals(".xlsx")):
        pass

    # Rename EXCEL Files
    if (ConcatenateMultiFiles == 0) and (ConcatenateHasHeader == 0) and (FileFormatIndicator == ".xlsx"):
        pass


def archive(LocalPath, ArchivePath, ArchiveDateFolderFlag, DestinationFilemask,
            SourceFilemask, SourceFilemask2, SourceFilemask3):
    if ArchiveDateFolderFlag == 0:
        if SourceFilemask is not None:
            for file in glob.glob(SourceFilemask):
                shutil.move(LocalPath + file, ArchivePath + file)
        else:
            pass
        if SourceFilemask2 is not None:
            for file in glob.glob(SourceFilemask2):
                shutil.move(LocalPath + file, ArchivePath + file)
        else:
            pass
        if SourceFilemask3 is not None:
            for file in glob.glob(SourceFilemask3):
                shutil.move(LocalPath + file, ArchivePath + file)
        else:
            pass
        if DestinationFilemask is not None:
            for file in glob.glob(DestinationFilemask + "*"):
                shutil.move(LocalPath + file, ArchivePath + file)
        else:
            pass
    else:
        if SourceFilemask is not None:
            for file in glob.glob(SourceFilemask):
                shutil.move(LocalPath + file, ArchivePath + "/" + str(datetime.now().month) + "/" +
                            str(datetime.now().day) + "/" + file)
        else:
            pass
        if SourceFilemask2 is not None:
            for file in glob.glob(SourceFilemask2):
                shutil.move(LocalPath + file, ArchivePath + "/" + str(datetime.now().month) + "/" +
                            str(datetime.now().day) + "/" + file)
        else:
            pass
        if SourceFilemask3 is not None:
            for file in glob.glob(SourceFilemask3):
                shutil.move(LocalPath + file, ArchivePath + "/" + str(datetime.now().month) + "/" +
                            str(datetime.now().day) + "/" + file)
        if DestinationFilemask is not None:
            for file in glob.glob(DestinationFilemask):
                shutil.move(LocalPath + file, ArchivePath + "/" + str(datetime.now().month) + "/" +
                            str(datetime.now().day) + "/" + file)
        else:
            pass


def ftp_get(SourceSFTPHost, SourceSFTPPort, SourceSFTPUserName, SourceSFTPPassword, LocalPath, SourceSFTPRemotePath,
            SourceFilemask, SourceFilemask2, SourceFilemask3):
    os.chdir(LocalPath)

    ftp = FTP(host=SourceSFTPHost, user=SourceSFTPUserName, passwd=SourceSFTPPassword)
    ftp.connect(host=SourceSFTPHost, port=SourceSFTPPort)
    ftp.cwd(SourceSFTPRemotePath)

    # loop through files ref: http://www.informit.com/articles/article.aspx?p=686162&seqNum=7
    if SourceFilemask is not None:
        for file in glob.glob(SourceFilemask):
            callback = open(file, "wb+")
            ftp.retrbinary('RETR ' + file, callback=callback)

    if SourceFilemask2 is not None:
        for file in glob.glob(SourceFilemask2):
            callback = open(file, "wb+")
            ftp.retrbinary('RETR ' + file, callback=callback)

    if SourceFilemask3 is not None:
        for file in glob.glob(SourceFilemask3):
            callback = open(file, "wb+")
            ftp.retrbinary('RETR ' + file, callback=callback)

    ftp.quit()

def ftp_put(LocalPath, DestinationFileMask, DestinationSFTPRemotPath, DestinationSFTPHost, DestinationSFTPUserName,
            DestinationSFTPPort, DestinationSFTPPassword, PGPRequired, SourceFilemask, SourceFilemask2, SourceFilemask3):
    os.chdir(LocalPath)

    if PGPRequired:
        pgp("O")
    else:

        ftp = FTP(host=DestinationSFTPHost, user=DestinationSFTPUserName, passwd=DestinationSFTPPassword)
        ftp.connect(host=DestinationSFTPHost, port=DestinationSFTPPort)
        ftp.cwd(DestinationSFTPRemotPath)

        if SourceFilemask is not None:
            for file in glob.glob(SourceFilemask):
                with open(file, 'r') as f:
                    ftp.storlines('STOR ' + file, f)

        if SourceFilemask2 is not None:
            for file in glob.glob(SourceFilemask2):
                with open(file, 'r') as f:
                    ftp.storlines('STOR ' + file, f)

        if SourceFilemask3 is not None:
            for file in glob.glob(SourceFilemask3):
                with open(file, 'r') as f:
                    ftp.storlines('STOR ' + file, f)

        if DestinationFileMask is not None:
            for file in glob.glob(DestinationFileMask):
                with open(file, 'r') as f:
                    ftp.storlines('STOR ' + file, f)

        ftp.quit()


def email_success(EmailRecipients, EmailCCReceipients, EmailSubject, EmailBody, EmailAttachFile, SourceFilemask,
                  SourceFilemask2, SourceFilemask3, DestinationFilemask):
    if EmailCCReceipients is not None:
        bcc = '; '.join(EmailCCReceipients)
    else:
        bcc = ''

    msg = MIMEMultipart()
    msg['From'] = "noreply@boonchapman.com"
    msg['To'] = '; '.join(EmailRecipients)
    msg['Subject'] = EmailSubject
    body = EmailBody + ". Processed at " + datetime.now()
    msg.attach(MIMEText(body, 'plain'))

    if EmailAttachFile == 1:
        if SourceFilemask is not None:
            for file1 in glob.glob(SourceFilemask):
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(file1, "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename={}'.format(file1))
                encoders.encode_base64(part)
                msg.attach(part)
                server = smtplib.SMTP("EXCHG10", port=25)
                server.sendmail(msg['From'], msg['To'] + bcc, msg.as_string())

        if SourceFilemask2 is not None:
            for file2 in glob.glob(SourceFilemask2):
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(file2, "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename={}'.format(file2))
                encoders.encode_base64(part)
                msg.attach(part)
                server = smtplib.SMTP("EXCHG10", port=25)
                server.sendmail(msg['From'], msg['To'] + bcc, msg.as_string())

        if SourceFilemask3 is not None:
            for file3 in glob.glob(SourceFilemask3):
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(file3, "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename={}'.format(file3))
                encoders.encode_base64(part)
                msg.attach(part)
                server = smtplib.SMTP("EXCHG10", port=25)
                server.sendmail(msg['From'], msg['To'] + bcc, msg.as_string())

        if DestinationFilemask is not None:
            for file4 in glob.glob(DestinationFilemask + "*"):
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(file4, "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename={}'.format(file4))
                encoders.encode_base64(part)
                msg.attach(part)
                server = smtplib.SMTP("EXCHG10", port=25)
                server.sendmail(msg['From'], msg['To'] + bcc, msg.as_string())

    elif EmailAttachFile == 0:
        server = smtplib.SMTP("EXCHG10", port=25)
        server.sendmail(msg['From'], msg['To'] + bcc, msg.as_string())

# def is_non_zero_file(LocalPath, ):
    # for i in os.path
    #   fpath = LocalPath +
    #   return os.path.isfile(fpath) and os.path.getsize(fpath) > 0


def outbound_sub(entity):
    cursor = cnxn.cursor()
    cursor.execute(
        "SELECT "
        "Entity"
        ",InboundOutbound"
        ",SourceSFTPHost"
        ",SourceSFTPPort"
        ",SourceSFTPUserName"   
        ",SourceSFTPPassword"   # 5
        ",SourceSFTPRemotePath"
        ",SourceFilemask"
        ",SourceFilemask2"
        ",SourceFilemask3"      
        ",RemoveFromSourceSFTP"     # 10
        ",LocalPath"
        ",DestinationPath"
        ",ArchiveFlag"
        ",ArchivePath"          
        ",ArchiveDateFolderFlag"    # 15
        ",PGPRequired"
        ",PGPKeyID"
        ",PGPKeyPassword"
        ",ConcatenateMultiFiles"    
        ",ConcatenateHasHeader"     # 20
        ",RunSunday"
        ",RunMonday"
        ",RunTuesday"
        ",RunWednesday"         
        ",RunThursday"      # 25
        ",RunFriday"
        ",RunSaturday"
        ",DayOfMonth"
        ",DateStampFile"      
        ",EmailNotification"       # 30
        ",EmailNoFile"
        ",EmailRecipients"
        ",EmailCCReceipients"
        ",EmailSubject"         
        ",EmailBody"        # 35
        ",EmailAttachFile"
        ",ForwardFlag"
        ",DestinationSFTPHost"
        ",DestinationSFTPPort"  
        ",DestinationSFTPUserName"  # 40
        ",DestinationSFTPPassword"
        ",DestinationSFTPRemotePath"
        ",DestinationRenameFlag"
        ",DestinationFilemask"  
        ",DestinationFileExtension" # 45
        ",FileFormatIndicator "
        "FROM TalendWorkspaceSQL2.dbo.Utility_FileHandling_Reference WHERE Entity = '{}'".format("Aliera")
    )
    outbound_local_list = cursor.fetchall()

    # looping through list of tuples
    for x in outbound_local_list:

        # set current working directory to LocalPath
        os.chdir(x[11])

        # (O) FTP get file
        if x[6] is not None:           # SourceSFTPRemotePath
            ftp_get(x[2], x[3], x[4], x[5], x[11], x[6], x[7], x[8], x[9])
        else:
            pass

        # Verify if file exists wrt Source file masks 1-3
        if glob.glob(x[7]) or glob.glob(x[8]) or glob.glob(x[9]):
            # Check if file is empty or not
            # is_non_zero_file(x[11], x[7], x[8], x[9])

            # (O) Modify--WORK ON THIS
            if x[43] == 1:
                modify(x[7], x[8], x[9], x[44], x[45], x[11], x[19], x[20], x[46], x[29], entity)
            else:
                pass

            # (O) PGP Encrypt--WORK ON THIS
            if x[16] is True:
                pgp(x[1], x[44], x[17], x[18], x[45])

            # if file is empty: #WORK ON THIS
            #     break
            # else:
                # Send Files to DestinationSFTP
                if x[37] == 1:
                    ftp_put(x[11], x[42], x[38], x[40], x[39], x[41], x[16], x[7], x[8], x[9])

                # Delete PGP files
                for f in glob.glob("*.pgp"):
                    os.remove(f)

                # (O) Archive Files
                if x[13] == 1:
                    archive(x[11], x[14], x[15], x[44], x[7], x[8], x[9])

                # Delete original files
                local_path = x[11]
                pattern = x[7] | x[8] | x[9]
                for f in os.listdir(local_path):
                    if re.search(pattern, f):
                        os.remove(os.path.join(local_path, f))

                # (O) Email notification w/o file attachment
                # ref: http://naelshiab.com/tutorial-send-email-python/
                if x[36] == 0 and x[30] == 1:
                    email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44])
                elif x[36] == 1 and x[30] == 1:
                    email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44])

        # file was not found so break out of loop & continue to next tuple
        else:
            break


def io_status():
    cursor = cnxn.cursor()
    cursor.execute(
        'SELECT [Entity], [InboundOutbound] '
        'FROM [TalendWorkspaceSQL2].[dbo].[Utility_FileHandling_Reference] UFR '
        'WHERE ((DATEPART(Weekday,GETDATE()) = 1 and RunSunday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 2 and RunMonday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 3 and RunTuesday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 4 and RunWednesday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 5 and RunThursday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 6 and RunFriday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 7 and RunSaturday = 1)'
        'OR (DATEPART(Day, GETDATE()) = DayOfMonth)'
        'OR (DATEPART(Day, GETDATE()) IN (SELECT * FROM TalendWorkspaceSQL2.dbo.CSVNumbersToTable(UFR.[DayOfMonth]))))')

    # returning a list of tuples
    results = cursor.fetchall()

    # unpacking/iterating through list of tuples
    for x in results:
        if x[1] == 'I':
            # inbound_sub(x[0])
            pass
        elif x[1] == 'O':
            outbound_sub(x[0])


if __name__ == "__main__":
    io_status()
