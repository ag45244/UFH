import pyodbc, os, re, shutil, glob
from ftplib import FTP
from datetime import datetime
from os import rename
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import subprocess
import logging

log_path = ''   # directory where logging should report to i.e. 'C:/Users/PythonLogging/UFH_log.txt'
logging.basicConfig(filename=log_path, level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')
logging.disable(logging.CRITICAL)
logging.debug('Start of program')

cnxn = pyodbc.connect('Driver={SQL Server};'
                      'Server=SQL2;'
                      'Database=ClaimsGeneral;'
                      'Trusted_Connection=yes;')


def ftp_get(SourceSFTPHost, SourceSFTPPort, SourceSFTPUserName, SourceSFTPPassword, LocalPath, SourceSFTPRemotePath,
            SourceFilemask, SourceFilemask2, SourceFilemask3, RemoveFromSourceSFTP):
    os.chdir(LocalPath)

    ftp = FTP(host=SourceSFTPHost, user=SourceSFTPUserName, passwd=SourceSFTPPassword)
    ftp.cwd(SourceSFTPRemotePath)
    ftp.retrlines('LIST')
    # loop through files ref:
    if (SourceFilemask is not None) and (SourceFilemask in x for x in ftp.nlst()):
        for file in list(x for x in ftp.nlst() if SourceFilemask in x):
            callback = open(file, "wb+")
            try:
                ftp.retrbinary('RETR ' + file, callback.write)
                # delete file from FTP
                if RemoveFromSourceSFTP is True:
                    ftp.delete(file)
                else:
                    pass
            except Exception as e:
                logging.debug('UFH failed to retrieve files from FTP: ' + str(e))

    if (SourceFilemask2 is not None) and (SourceFilemask2 in x for x in ftp.nlst()):
        for file in list(x for x in ftp.nlst() if SourceFilemask2 in x):
            callback = open(file, "wb+")
            try:
                ftp.retrbinary('RETR ' + file, callback.write)
                # delete file from FTP
                if RemoveFromSourceSFTP is True:
                    ftp.delete(file)
                else:
                    pass
            except Exception as e:
                logging.debug('UFH failed to retrieve files from FTP: ' + str(e))

    if (SourceFilemask3 is not None) and (SourceFilemask3 in x for x in ftp.nlst()):
        for file in list(x for x in ftp.nlst() if SourceFilemask3 in x):
            callback = open(file, "wb+")
            try:
                ftp.retrbinary('RETR ' + file, callback.write)
                # delete file from FTP
                if RemoveFromSourceSFTP is True:
                    ftp.delete(file)
                else:
                    pass
            except Exception as e:
                logging.debug('UFH failed to retrieve files from FTP: ' + str(e))

    ftp.quit()


def ftp_put(LocalPath, DestinationFileMask, DestinationSFTPRemotPath, DestinationSFTPHost, DestinationSFTPUserName,
            DestinationSFTPPort, DestinationSFTPPassword, PGPRequired, SourceFilemask,
            SourceFilemask2, SourceFilemask3, File_List):
    os.chdir(LocalPath)
    if PGPRequired:
        pgp("O")

    ftp = FTP(host=DestinationSFTPHost, user=DestinationSFTPUserName, passwd=DestinationSFTPPassword)
    ftp.cwd(DestinationSFTPRemotPath)

    if SourceFilemask is not None:
        for file in glob.glob(SourceFilemask + "*"):
            with open(file, 'r') as f:
                ftp.storlines('STOR ' + file, f)
    if SourceFilemask2 is not None:
        for file in glob.glob(SourceFilemask2 + "*"):
            with open(file, 'r') as f:
                ftp.storlines('STOR ' + file, f)
    if SourceFilemask3 is not None:
        for file in glob.glob(SourceFilemask3 + "*"):
            with open(file, 'r') as f:
                ftp.storlines('STOR ' + file, f)
    if DestinationFileMask is not None:
        for file in glob.glob(DestinationFileMask + "*"):
            with open(file, 'r') as f:
                ftp.storlines('STOR ' + file, f)

    for file in File_List:
        with open(file, 'r') as f:
            ftp.storlines('STOR ' + file, f)

    ftp.quit()


def modify(File_List, DestinationFilemask,
           DestinationFileExtension, LocalPath, ConcatenateMultiFiles, ConcatenateHasHeader,
           FileFormatIndicator, DateStampFile, entity):
    # Rename files with date time stamp, and numeric.sequence()
    if (ConcatenateMultiFiles == 0) and (ConcatenateHasHeader == 0) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi") or
            (FileFormatIndicator == ".xlsx")):

        for x in range(1, len(File_List) + 1):
            if DateStampFile is True:
                rename(LocalPath + '/' + File_List[x - 1], LocalPath + '/' + File_List[x - 1].replace('.',
                                                                                                      str(
                                                                                                          datetime.now().year) + str(
                                                                                                          datetime.now().month) + str(
                                                                                                          datetime.now().day) + '_' + str(
                                                                                                          x) + '_.'))
            elif DateStampFile is False:
                rename(LocalPath + '/' + File_List[x - 1],
                       LocalPath + '/' + File_List[x - 1].replace('.', '_' + str(x) + '_.'))

    output_filename = ''
    if DateStampFile is True:
        output_filename = (DestinationFilemask + '_' + str(datetime.now().year) + str(datetime.now().month) +
                           str(datetime.now().day) + '_' + DestinationFileExtension)
    elif DateStampFile is False:
        output_filename = (DestinationFilemask + DestinationFileExtension)

    # Rename Concatenated TEXT/CSV Files INCLUDE Header--Normal appending of files
    if (ConcatenateMultiFiles == 1) and (ConcatenateHasHeader == 1) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi")):

        with open(LocalPath + '/' + output_filename, 'w', newline='\n') as outfile:
            for fname in File_List:
                with open(fname) as infile:
                    for line in infile:
                        outfile.write(line)

    # Rename Concatenated TEXT/CSV Files EXCLUDE Header--will have to use readline method to somehow exclude first line
    if (ConcatenateMultiFiles == 1) and (ConcatenateHasHeader == 0) and (
            (FileFormatIndicator == ".txt") or (FileFormatIndicator == ".csv") or (FileFormatIndicator == ".edi")):

        with open(LocalPath + '/' + output_filename, 'w') as outfile:
            for fname in File_List:
                with open(fname) as infile:
                    # header = infile.readlines()[:1]
                    for line in infile.readlines()[1:]:
                        outfile.write(line)
                    outfile.write('\n')

    # Rename Concatenated EXCEL Files
    if (ConcatenateMultiFiles == 1) and (ConcatenateHasHeader == 0) and (FileFormatIndicator == ".xlsx"):
        # read excel files in
        excels = [pd.ExcelFile(name) for name in File_List]

        # convert to data frames
        frames = [x.parse(x.sheet_names[0], header=None, index_col=None) for x in excels]

        # delete the first row for all frames except the first
        # i.e. remove the header row -- assumes it's the first
        frames[1:] = [df[1:] for df in frames[1:]]

        # concatenate them..
        combined = pd.concat(frames)

        # write it out
        combined.to_excel(output_filename, header=False, index=False)


def email_success(EmailRecipients, EmailCCReceipients, EmailSubject, EmailBody, EmailAttachFile, SourceFilemask,
                  SourceFilemask2, SourceFilemask3, DestinationFilemask, file_list):
    if EmailCCReceipients is not None:
        bcc = '; '.join(EmailCCReceipients)
    else:
        bcc = ''

    msg = MIMEMultipart()
    msg['From'] = ""    # email address omitted
    msg['To'] = '; '.join(EmailRecipients)
    msg['Subject'] = EmailSubject
    body = EmailBody + ". Processed at " + str(datetime.now())
    msg.attach(MIMEText(body, 'plain'))

    if EmailAttachFile is True:
        for file1 in file_list:
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(file1, "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename={}'.format(file1))
            encoders.encode_base64(part)
            msg.attach(part)

            server = smtplib.SMTP("smtp-mail.outlook.com", port=587)
            server.starttls()
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            server.quit()

    elif EmailAttachFile is False:
        server = smtplib.SMTP("smtp.gmail.com", port=587)
        server.starttls()
        server.login("", "")    # Email and password omitted
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()


def inbound_sub(entity):
    cursor = cnxn.cursor()
    cursor.execute(
        "SELECT "
        "Entity"
        ",InboundOutbound"
        ",SourceSFTPHost"
        ",SourceSFTPPort"
        ",SourceSFTPUserName"
        ",SourceSFTPPassword"  # 5
        ",SourceSFTPRemotePath"
        ",SourceFilemask"
        ",SourceFilemask2"
        ",SourceFilemask3"
        ",RemoveFromSourceSFTP"  # 10
        ",LocalPath"
        ",DestinationPath"
        ",ArchiveFlag"
        ",ArchivePath"
        ",ArchiveDateFolderFlag"  # 15
        ",PGPRequired"
        ",PGPKeyID"
        ",PGPKeyPassword"
        ",ConcatenateMultiFiles"
        ",ConcatenateHasHeader"  # 20
        ",RunSunday"
        ",RunMonday"
        ",RunTuesday"
        ",RunWednesday"
        ",RunThursday"  # 25
        ",RunFriday"
        ",RunSaturday"
        ",DayOfMonth"
        ",DateStampFile"
        ",EmailNotification"  # 30
        ",EmailNoFile"
        ",EmailRecipients"
        ",EmailCCReceipients"
        ",EmailSubject"
        ",EmailBody"  # 35
        ",EmailAttachFile"
        ",ForwardFlag"
        ",DestinationSFTPHost"
        ",DestinationSFTPPort"
        ",DestinationSFTPUserName"  # 40
        ",DestinationSFTPPassword"
        ",DestinationSFTPRemotePath"
        ",DestinationRenameFlag"
        ",DestinationFilemask"
        ",DestinationFileExtension"  # 45
        ",FileFormatIndicator "
        "FROM TalendWorkspaceSQL2.dbo.Utility_FileHandling_Reference WHERE Entity = '{}'".format(entity)
    )
    outbound_local_list = cursor.fetchall()
    # looping through list of tuples
    for x in outbound_local_list:
        # set current working directory to LocalPath
        os.chdir(x[11])

        # create file name list (DRY)
        file_list = list()
        [file_list.append(file) for file in glob.glob(str(x[7]) + "*") if x[7] is not None]
        [file_list.append(file) for file in glob.glob(str(x[8]) + "*") if x[8] is not None]
        [file_list.append(file) for file in glob.glob(str(x[9]) + "*") if x[9] is not None]

        # (O) FTP get file
        if x[42] is not None:  # DestinationSourceSFTPRemotePath
            ftp_get(x[2], x[3], x[4], x[5], x[11], x[6], x[7], x[8], x[9], x[10])
        else:
            pass

        # (O) PGP Decrypt
        if x[16] is True:
            pgp(x[1], x[44], x[17], x[18], x[45])

        # Delete PGP files
        for f in glob.glob("*.pgp"):
            os.remove(f)

        # Verify if file exists wrt Source file masks 1-3
        if (glob.glob(x[7] + "*") or glob.glob(x[8] + "*") or glob.glob(x[9] + "*")) is not None:

            # Check if file is empty or not
            for file in (x for x in glob.glob(x[7] + "*") if x[7] in x):
                if os.path.exists(x[11] + file) and os.path.getsize(x[11] + file) > 0:
                    pass
                else:
                    break
        # (O) Modify
        if x[43] == 1:
            modify(file_list, x[44], x[45], x[11], x[19], x[20], x[46], x[29], entity)
        else:
            pass

        # Send Files to DestinationSFTP
        if x[37] == 1:
            ftp_put(x[11], x[44], x[42], x[38], x[40], x[39], x[41], x[16], x[7], x[8], x[9], file_list)
        else:
            pass

        # (O) Archive Files
        if x[13] == 1:
            archive(x[11], x[14], x[15], x[44], x[7], x[8], x[9], file_list)

        # Delete original files
        local_path = x[11]
        pattern = x[7] | x[8] | x[9]
        for f in os.listdir(local_path):
            if re.search(pattern, f):
                os.remove(os.path.join(local_path, f))

        # (O) Email notification w/o file attachment
        if x[36] == 0 and x[30] == 1:
            email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44], file_list)
        elif x[36] == 1 and x[30] == 1:
            email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44])

        # file was not found so break & continue to next tuple
        else:
            break


def outbound_sub(entity):
    cursor = cnxn.cursor()
    cursor.execute(
        "SELECT "
        "Entity"
        ",InboundOutbound"
        ",SourceSFTPHost"
        ",SourceSFTPPort"
        ",SourceSFTPUserName"
        ",SourceSFTPPassword"  # 5
        ",SourceSFTPRemotePath"
        ",SourceFilemask"
        ",SourceFilemask2"
        ",SourceFilemask3"
        ",RemoveFromSourceSFTP"  # 10
        ",LocalPath"
        ",DestinationPath"
        ",ArchiveFlag"
        ",ArchivePath"
        ",ArchiveDateFolderFlag"  # 15
        ",PGPRequired"
        ",PGPKeyID"
        ",PGPKeyPassword"
        ",ConcatenateMultiFiles"
        ",ConcatenateHasHeader"  # 20
        ",RunSunday"
        ",RunMonday"
        ",RunTuesday"
        ",RunWednesday"
        ",RunThursday"  # 25
        ",RunFriday"
        ",RunSaturday"
        ",DayOfMonth"
        ",DateStampFile"
        ",EmailNotification"  # 30
        ",EmailNoFile"
        ",EmailRecipients"
        ",EmailCCReceipients"
        ",EmailSubject"
        ",EmailBody"  # 35
        ",EmailAttachFile"
        ",ForwardFlag"
        ",DestinationSFTPHost"
        ",DestinationSFTPPort"
        ",DestinationSFTPUserName"  # 40
        ",DestinationSFTPPassword"
        ",DestinationSFTPRemotePath"
        ",DestinationRenameFlag"
        ",DestinationFilemask"
        ",DestinationFileExtension"  # 45
        ",FileFormatIndicator "
        "FROM TalendWorkspaceSQL2.dbo.Utility_FileHandling_Reference WHERE Entity = '{}'".format(entity)
    )
    outbound_local_list = cursor.fetchall()
    # looping through list of tuples
    for x in outbound_local_list:
        # set current working directory to LocalPath
        os.chdir(x[11])

        # create file name list (DRY)
        file_list = list()
        [file_list.append(file) for file in glob.glob(str(x[7]) + "*") if x[7] is not None]
        [file_list.append(file) for file in glob.glob(str(x[8]) + "*") if x[8] is not None]
        [file_list.append(file) for file in glob.glob(str(x[9]) + "*") if x[9] is not None]

        # (O) FTP get file
        if x[6] is not None:  # SourceSFTPRemotePath
            ftp_get(x[2], x[3], x[4], x[5], x[11], x[6], x[7], x[8], x[9], x[10])
        else:
            pass

        # Verify if file exists wrt Source file masks 1-3
        if (glob.glob(x[7] + "*") or glob.glob(x[8] + "*") or glob.glob(x[9] + "*")) is not None:

            # Check if file is empty or not
            for file in (x for x in glob.glob(x[7] + "*") if x[7] in x):
                if os.path.exists(x[11] + file) and os.path.getsize(x[11] + file) > 0:
                    pass
                else:
                    break
        # (O) Modify
        if x[43] == 1:
            modify(file_list, x[44], x[45], x[11], x[19], x[20], x[46], x[29], entity)
        else:
            pass

        # (O) PGP Encrypt
        if x[16] is True:
            pgp(x[1], x[44], x[17], x[18], x[45])

        # Send Files to DestinationSFTP
        if x[37] == 1:
            ftp_put(x[11], x[44], x[42], x[38], x[40], x[39], x[41], x[16], x[7], x[8], x[9], file_list)
        else:
            pass
        # Delete PGP files
        for f in glob.glob("*.pgp"):
            os.remove(f)

        # (O) Archive Files
        if x[13] == 1:
            archive(x[11], x[14], x[15], x[44], x[7], x[8], x[9], file_list)

        # Delete original files
        local_path = x[11]
        pattern = x[7] | x[8] | x[9]
        for f in os.listdir(local_path):
            if re.search(pattern, f):
                os.remove(os.path.join(local_path, f))

        # (O) Email notification w/o file attachment
        if x[36] == 0 and x[30] == 1:
            email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44], file_list)
        elif x[36] == 1 and x[30] == 1:
            email_success([32], x[33], x[34], x[35], x[36], x[7], x[8], x[9], x[44])

        # file was not found so break out of loop & continue to next tuple
        else:
            break


def io_status(entity):
    cursor = cnxn.cursor()
    cursor.execute(
        'SELECT [Entity], [InboundOutbound] '
        'FROM [TalendWorkspaceSQL2].[dbo].[Utility_FileHandling_Reference] UFR '
        'WHERE (UFR.Entity={}) AND ((DATEPART(Weekday,GETDATE()) = 1 and RunSunday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 2 and RunMonday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 3 and RunTuesday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 4 and RunWednesday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 5 and RunThursday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 6 and RunFriday = 1) '
        'OR (DATEPART(Weekday,GETDATE()) = 7 and RunSaturday = 1)'
        'OR (DATEPART(Day, GETDATE()) = DayOfMonth)'
        'OR (DATEPART(Day, GETDATE()) IN (SELECT * '
        'FROM TalendWorkspaceSQL2.dbo.CSVNumbersToTable(UFR.[DayOfMonth]))))'.format(entity))

    # returning a list of tuples
    results = cursor.fetchall()

    # unpacking/iterating through list of tuples
    for x in results:
        if x[1] == 'I':
            inbound_sub(x[0])
            pass
        elif x[1] == 'O':
            outbound_sub(x[0])


if __name__ == "__main__":
    io_status()
