#########################################################
################ -- EMAIL - SENDING -- ##################
#########################################################


from re import sub
import smtplib
import smtplib, ssl
from email.message import EmailMessage
import mimetypes
import os.path

import variables
import rsa

import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

def startBrowser(path):
    if not os.path.exists(path):
        os.mkdir(path)
    chrom_options_ = webdriver.ChromeOptions()
    prefer_ = {'download.default_directory': path,
                'profile.default_content_settings.popups': 0,
                'directory_upgrade': True}

    chrom_options_.add_experimental_option('prefs',prefer_)
    chrom_options_.add_argument("user-data-dir={}".format(f"{path}\\chrome"))
    chrom_options_.add_argument("start-maximized")
    #chrom_options_.add_argument("--headless")  

    desired_capabilities = DesiredCapabilities.CHROME.copy()
    desired_capabilities['acceptInsecureCerts'] = True
    browser = None
    try:
        exctbl_path = chromedriver_autoinstaller.install(True)
        with open("chromeDriver.txt", "w") as text_file:
            print(f"{exctbl_path}", file=text_file)
            text_file.close()

        browser = webdriver.Chrome( executable_path = exctbl_path, desired_capabilities=desired_capabilities, options=chrom_options_)
    except:
        print("<<<<<<no se pudo instalar chrome driver intentando iniciar con anterior version>>>>>>")
        try:
            with open('chromeDriver.txt', 'r') as text_file:
                exctbl_path = text_file.read().replace('\n', '')
            browser = webdriver.Chrome( executable_path = exctbl_path, desired_capabilities=desired_capabilities, options=chrom_options_)
        except:
            print("<<<<<<no se pudo instalar chrome driver intentando iniciar con version de Servicios Comerciales>>>>>>")
            
    browser.get("chrome://settings/")
    browser.execute_script("chrome.settingsPrivate.setDefaultZoom(1.0);")
    return browser

def sendEmail (recipients, subject, from_name = 'Robot Suspensiones', attachment_files = [], cc=None):
    port = 465  # For SSL
    smtp_server = "mail.epm.com.co"
    encryptedFrom = b'\x1b\xa0\x1e\xd4N\xc2M\xbc#\xff\x9bRv?,!\x95\x1e\x95\x0euy\\;3\x87\\\t\nz\xfc\x88\xc9\x8bS\xa9\xa5\xa4\xd2\xcc\xdc\xce]\x82t\x19kG\x13\xadL\x1bwl\x91\x9b-\xaa\xf5\xdaF(9\x05qf2\xea\x9b\xf3PR\xe6\xef7\xc76\x99\xe2_\xe7\x8c\xd0\xdb\xd3\xca\x97/JI\xf0\xb7\x12MmAl\xd4Ur\xf1#-\x13.\n\xa6\xdd\x8dD\x8f4\x1bTX\r\xcaA\x18\xfa\xc4#"\xac\x0b\xafTy\x0b\xe6o\x18Y\xd9\xdf\xc6]\x97x\x17c\xd4\x95Z\x9f\x82\xb3\xe7\xf1\xbdqz\'\x0cT\xb0\xa3\xaaK&\xd0t\xf0\xa39\xdc\x8a6n\xda\x00E\xd1W\n\x13x\x0e\xde\x92$\x81\x9a\n\xae)>O+lHv\xf1zK\x00\tu\xb3\xce\x85\x86x\xde*S\xde\x8e\xab\xd6p\xb59\x0cqoZ\x86\xdd\xc6/\xa9p0\xbc\x10h\x82\xfd\x1a\x83\xdb\x18\xc9\xca|1\xdd\x8a\x10\xf8\xa1\\\xd1\xd5\xc7\xb8GX\xd3\x02u\xbf\x16\xd0e'
    encryptedPsw = b'B\xe8\x94\r\x1bb\xd6\x87\x8d\xf0\x96\xd9\x10\x95Q\xb7\'\x81\x17eW\xf1L\x02\x0b\xc0\xa6\xca\x16\x1b~I\xbev\xd8\x0cU\xcaJ\xb1I\x80\x04o\xffd\xe8l1\xfa2\xeb\xe9\xcd\x807\xaaVlz\x1e~08\xb0\x80\x0f\\\x88%T\x9cK\xa2\xe6\xeb\x08\x84b4CZ\x80\xfe-\x9d\xee\'!\x95\xe8:@\xb9\x8f\xd8\xf9\xb3\x1f&\xee\xd0K-\xc2\'\xdc\xd47\x13I,\xc5n;\xf9-\xb1\x82\xf7\x8b\xd6V\x19{\n\xf61\x86\x0cn\xb0\xd9\xf9/.T\xc5\x07\x90\x1e\xd7\xa5\x1f\xa3\xcc\xb3z\x0ed\xe5\n\xcc\x1a\xda2\x8b\x1e\x1cX\xa5\xd6CN\xcd\x9f\xad\xb4\x93\xd8\xf4\x94V\xc5\nok\x992\x88\xdbk\xd4P\x82d|\x97\x89\x0cO\\\xbb\xf8\x94\x8cv,+\x82PFL\xa6\xc6Vn\xed\xb2\xeeB\xcd^\xfa\x8c\x15 \x8be=\xde\x02\x9aQ#\xef\x93\x19\xce\x82\xf7E\xdf\x1d\xdb|\xc6\x98",\t\x06C\x9fC\x86\\\xa4\x99\x15\t\xf0\x175r\x11'
    encrypted_login = b'd\x8d\xe1\xe6l\x18\x80\x8a\x97\xfc\xc3\xff\t\xf3\x1cZ\xe5\xcdN\x83\x00\xf5\xcc\x11\xbdUD\x12\xcd\x8aG\x90\xec 5F\x02~\x1d\x16\xdf\xf5\x1c\x04\x86\xdc\xcf\xac\x899\xd5\x00\xabx\xd6\x02\xd7\xd6\xceHF\xd9}\x0c \x8cLWd\xe5\xd5\x8ca\x17\xc9\xb5\xa0\xae\xae\x0b\xa2\x8e/\x0c \x86p\t\t\xc2\xf1kA8\x08\x06\xe4\x18R\x0b\x18\x04(\xf7\x9e\x08\xde\xd3hm\xb7$R\xb4a4\'\x0e\x10\xf6\x81\xf8\xba\xf2\x0c\x82$\xa0\x80>\x96\xa1\xb3\x1c\tP\x04&\x9b.\xf0\rD\x9f\x8f\xbcv,{@ZY\xd2\xae\xa2pT\xcd\x8ab\xc3\xefi\x88a\xd5\xa0P@\x04\x08 \xce\xb7\x9b"[\xd8\xf2.\xaf\x11\xce \xb89\xc8U\xe7\x97l\\\xfd\x03Gx\xa3\xd4|\xcfZ\x92\xf2\x86\x015#J\x84\x99\x9f\x98\x9d\xcd\x07D(1\xec\xba\xcaI\x9d8\x8c\xcd\xff\xbc\x8e;\x0ePo\xfaq\xdc&\xe2\xc5\xeczQ\x18\xee.\x9f\xe1\x93>\x92\x98\x1e\x15j\x10H'

    with open('key.pkey', 'rb') as privatefile:
        keydata = privatefile.read()
        privkey = rsa.PrivateKey.load_pkcs1(keydata, 'DER')
    
    password = rsa.decrypt(encryptedPsw, privkey).decode()
    sender_email = rsa.decrypt(encryptedFrom, privkey).decode()
    login = rsa.decrypt(encrypted_login, privkey).decode()
    message = EmailMessage()
    message["Subject"] = f'{subject} - version: {variables.version}'
    message["From"] = f'{from_name} <{sender_email}>'
    message["To"] = ", ".join(recipients)
    if cc:
        message['Cc'] = ", ".join(cc)

    # set the plain text body
    message.set_content("""Adjunto encontrará el análisis de critica generados el dia de hoy. 
    Esta es una notificacion automática no responda este correo. 
    Si tiene alguna duda comuniquese con William Castrillón""")

    # set an alternative html body
    message.add_alternative("""<html xmlns=http://www.w3.org/1999/xhtml xmlns:o=urn:schemas-microsoft-com:office:office xmlns:v=urn:schemas-microsoft-com:vml><head><!--[if gte mso 9]><xml><o:officedocumentsettings><o:allowpng><o:pixelsperinch>96</o:pixelsperinch></o:officedocumentsettings></xml><![endif]--><meta content="text/html; charset=utf-8"http-equiv=Content-Type><meta content="width=device-width"name=viewport><!--[if !mso]><!--><meta content="IE=edge"http-equiv=X-UA-Compatible><!--<![endif]--><title></title><!--[if !mso]><!--><!--<![endif]--><style>body{margin:0;padding:0}table,td,tr{vertical-align:top;border-collapse:collapse}*{line-height:inherit}a[x-apple-data-detectors=true]{color:inherit!important;text-decoration:none!important}</style><style id=media-query>@media (max-width:670px){.block-grid,.col{min-width:320px!important;max-width:100%!important;display:block!important}.block-grid{width:100%!important}.col{width:100%!important}.col_cont{margin:0 auto}img.fullwidth,img.fullwidthOnMobile{max-width:100%!important}.no-stack .col{min-width:0!important;display:table-cell!important}.no-stack.two-up .col{width:50%!important}.no-stack .col.num2{width:16.6%!important}.no-stack .col.num3{width:25%!important}.no-stack .col.num4{width:33%!important}.no-stack .col.num5{width:41.6%!important}.no-stack .col.num6{width:50%!important}.no-stack .col.num7{width:58.3%!important}.no-stack .col.num8{width:66.6%!important}.no-stack .col.num9{width:75%!important}.no-stack .col.num10{width:83.3%!important}.video-block{max-width:none!important}.mobile_hide{min-height:0;max-height:0;max-width:0;display:none;overflow:hidden;font-size:0}.desktop_hide{display:block!important;max-height:none!important}}</style><style id=icon-media-query>@media (max-width:670px){.icons-inner{text-align:center}.icons-inner td{margin:0 auto}}</style><body class=clean-body style=margin:0;padding:0;-webkit-text-size-adjust:100%;background-color:#3d1554><!--[if IE]><div class=ie-browser><![endif]--><table cellpadding=0 cellspacing=0 role=presentation style=table-layout:fixed;vertical-align:top;min-width:320px;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;background-color:#3d1554;width:100% valign=top class=nl-container width=100% bgcolor=#3d1554><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top valign=top><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=background-color:#3d1554 align=center><![endif]--><div style=background-color:#57366e><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:#57366e><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px>
    <tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:5px;padding-bottom:5px><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:650px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:5px;padding-bottom:5px;padding-right:0;padding-left:0"><!--<![endif]--><table cellpadding=0 cellspacing=0 role=presentation style=table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;min-width:100%;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100% valign=top class=divider width=100% border=0><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top;min-width:100%;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%;padding-top:10px;padding-right:10px;padding-bottom:10px;padding-left:10px valign=top class=divider_inner><table cellpadding=0 cellspacing=0 role=presentation style="table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;border-top:0 solid transparent;height:0;width:100%"valign=top class=divider_content width=100% border=0 align=center height=0><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100% valign=top height=0><span></span></table></table><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:#57366e><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class="block-grid mixed-two-up"><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:#57366e><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:162px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=162><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:0;padding-bottom:0>
    <![endif]--><div style=display:table-cell;vertical-align:top;max-width:320px;min-width:162px;width:162px class="col num3"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:0;padding-bottom:0;padding-right:0;padding-left:0"><!--<![endif]--><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:487px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=487><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:15px;padding-bottom:10px><![endif]--><div style=display:table-cell;vertical-align:top;max-width:320px;min-width:486px;width:487px class="col num9"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:15px;padding-bottom:10px;padding-right:0;padding-left:0"><!--<![endif]--><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:10px;padding-left:10px;padding-top:10px;padding-bottom:10px;font-family:Arial,sans-serif><![endif]--><!--[if mso]><![endif]--><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:#57366e><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:#57366e><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:5px;padding-bottom:5px><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:650px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:5px;padding-bottom:5px;padding-right:0;padding-left:0"><!--<![endif]-->
    <table cellpadding=0 cellspacing=0 role=presentation style=table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;min-width:100%;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100% valign=top class=divider width=100% border=0><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top;min-width:100%;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%;padding-top:10px;padding-right:10px;padding-bottom:10px;padding-left:10px valign=top class=divider_inner><table cellpadding=0 cellspacing=0 role=presentation style="table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;border-top:0 solid transparent;height:0;width:100%"valign=top class=divider_content width=100% border=0 align=center height=0><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top;-ms-text-size-adjust:100%;-webkit-text-size-adjust:100% valign=top height=0><span></span></table></table><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:transparent><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:transparent><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:35px;padding-bottom:0><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:650px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:35px;padding-bottom:0;padding-right:0;padding-left:0"><!--<![endif]--><div style=padding-right:0;padding-left:0 class="autowidth center img-container"align=center><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr style=line-height:0><td style=padding-right:0;padding-left:0 align=center><![endif]--> <img align=center border=0 class="autowidth center"src=cid:header style=text-decoration:none;-ms-interpolation-mode:bicubic;height:auto;border:0;width:100%;max-width:650px;display:block width=650><!--[if mso]><![endif]--></div>
    <!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:transparent><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:transparent><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:4px solid #57366E;border-bottom:0 solid transparent;border-right:4px solid #57366E"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:55px;padding-bottom:60px><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:642px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:4px solid #57366E;border-bottom:0 solid transparent;border-right:4px solid #57366E;padding-top:55px;padding-bottom:60px;padding-right:0;padding-left:0"><!--<![endif]--><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:10px;padding-left:10px;padding-top:10px;padding-bottom:10px;font-family:Arial,sans-serif><![endif]--><div style=color:#fbd711;font-family:Poppins,Arial,Helvetica,sans-serif;line-height:1.2;padding-top:10px;padding-right:10px;padding-bottom:10px;padding-left:10px><div style=line-height:1.2;font-size:12px;color:#fbd711;font-family:Poppins,Arial,Helvetica,sans-serif;mso-line-height-alt:14px class=txtTinyMce-wrapper><p style=margin:0;font-size:14px;line-height:1.2;word-break:break-word;text-align:center;mso-line-height-alt:17px;margin-top:0;margin-bottom:0><strong><span style=font-size:30px>Hola,</span></strong></div></div><!--[if mso]><![endif]--><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:50px;padding-left:50px;padding-top:10px;padding-bottom:10px;font-family:Arial,sans-serif><![endif]--><div style=color:#fff;font-family:Poppins,Arial,Helvetica,sans-serif;line-height:1.2;padding-top:10px;padding-right:50px;padding-bottom:10px;padding-left:50px><div style=line-height:1.2;font-size:12px;color:#fff;font-family:Poppins,Arial,Helvetica,sans-serif;mso-line-height-alt:14px class=txtTinyMce-wrapper><p style=margin:0;font-size:28px;line-height:1.2;word-break:break-word;text-align:center;mso-line-height-alt:34px;margin-top:0;margin-bottom:0><span style=font-size:28px>Adjunto encontrará las SDS que se encuentran pendientes el día de hoy.</span>
    <p style=margin:0;font-size:14px;line-height:1.2;word-break:break-word;text-align:center;mso-line-height-alt:17px;margin-top:0;margin-bottom:0><p style=margin:0;font-size:28px;line-height:1.2;word-break:break-word;text-align:center;mso-line-height-alt:34px;margin-top:0;margin-bottom:0><span style=font-size:28px><em><span style=font-size:22px;color:#fbd711><strong>Esta es una notificacion automática no responda este correo. Si tiene alguna duda comuniquese con Manuel Montero</strong></span></em></span></div></div><!--[if mso]><![endif]--><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:transparent><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:transparent><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:0;padding-bottom:0><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:650px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:0;padding-bottom:0;padding-right:0;padding-left:0"><!--<![endif]--><div style=padding-right:0;padding-left:0 class="autowidth center img-container"align=center><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr style=line-height:0><td style=padding-right:0;padding-left:0 align=center><![endif]--> <img align=center border=0 class="autowidth center"src=cid:bottom style=text-decoration:none;-ms-interpolation-mode:bicubic;height:auto;border:0;width:100%;max-width:650px;display:block width=650><!--[if mso]><![endif]--></div><!--[if mso]><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:10px;padding-left:10px;padding-top:15px;padding-bottom:20px;font-family:Arial,sans-serif><![endif]--><div style=color:#b0a7b7;font-family:Poppins,Arial,Helvetica,sans-serif;line-height:1.5;padding-top:15px;padding-right:10px;padding-bottom:20px;padding-left:10px><div style=line-height:1.5;font-size:12px;color:#b0a7b7;font-family:Poppins,Arial,Helvetica,sans-serif;mso-line-height-alt:18px class=txtTinyMce-wrapper></div></div>
    <!--[if mso]><![endif]--><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><div style=background-color:transparent><div style="min-width:320px;max-width:650px;overflow-wrap:break-word;word-wrap:break-word;word-break:break-word;Margin:0 auto;background-color:transparent"class=block-grid><div style=border-collapse:collapse;display:table;width:100%;background-color:transparent><!--[if (mso)|(IE)]><table cellpadding=0 cellspacing=0 width=100% border=0 style=background-color:transparent><tr><td align=center><table cellpadding=0 cellspacing=0 border=0 style=width:650px><tr style=background-color:transparent class=layout-full-width><![endif]--><!--[if (mso)|(IE)]><td style="background-color:transparent;width:650px;border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent"valign=top align=center width=650><table cellpadding=0 cellspacing=0 width=100% border=0><tr><td style=padding-right:0;padding-left:0;padding-top:5px;padding-bottom:5px><![endif]--><div style=min-width:320px;max-width:650px;display:table-cell;vertical-align:top;width:650px class="col num12"><div style=width:100%!important class=col_cont><!--[if (!mso)&(!IE)]><!--><div style="border-top:0 solid transparent;border-left:0 solid transparent;border-bottom:0 solid transparent;border-right:0 solid transparent;padding-top:5px;padding-bottom:5px;padding-right:0;padding-left:0"><!--<![endif]--><table cellpadding=0 cellspacing=0 role=presentation style=table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0 valign=top width=100%><tr style=vertical-align:top valign=top><td style=word-break:break-word;vertical-align:top;padding-top:5px;padding-right:0;padding-bottom:5px;padding-left:0;text-align:center valign=top align=center><!--[if vml]><table cellpadding=0 cellspacing=0 role=presentation style=display:inline-block;padding-left:0;padding-right:0;mso-table-lspace:0;mso-table-rspace:0 align=left><![endif]--><!--[if !vml]><!--><table cellpadding=0 cellspacing=0 role=presentation style=table-layout:fixed;vertical-align:top;border-spacing:0;border-collapse:collapse;mso-table-lspace:0;mso-table-rspace:0;display:inline-block;margin-right:-4px;padding-left:0;padding-right:0 valign=top class=icons-inner><!--<![endif]--><tr style=vertical-align:top valign=top></table></table><!--[if (!mso)&(!IE)]><!--></div><!--<![endif]--></div></div><!--[if (mso)|(IE)]><![endif]--><!--[if (mso)|(IE)]><![endif]--></div></div></div><!--[if (mso)|(IE)]><![endif]--></table><!--[if (IE)]><![endif]-->""", subtype='html')

    # now open the image and attach it to the email
    with open('content-top_2.png', 'rb') as img:
        # know the Content-Type of the image
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        # attach it
        message.get_payload()[1].add_related(img.read(), 
                                            maintype=maintype, 
                                            subtype=subtype, 
                                            cid="header")
    
    with open('content-bottom_1.png', 'rb') as img:
        # know the Content-Type of the image
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        # attach it
        message.get_payload()[1].add_related(img.read(), 
                                            maintype=maintype, 
                                            subtype=subtype, 
                                            cid="bottom")

    if len(attachment_files) > 0:
        # Open attachment file in binary mode
        for attachment_file in attachment_files:
            with open(attachment_file, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                maintype, subtype = mimetypes.guess_type(attachment.name)[0].split('/')
                print(maintype)
                print(subtype)
                message.add_attachment( attachment.read(), maintype=maintype, subtype=subtype, filename=attachment_file)
    
    # Create secure connection with server and send email
    context = ssl._create_unverified_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.login(login, password)
        server.send_message(message)
    return True

#################################################################################
############################### CHECK FOR UPDATES ###############################
#################################################################################
    
import imaplib
import email
from email.header import decode_header
import webbrowser
import os


# busca actualizaciones basado en correo, la version actual del programa viene en el body del email, el subject sirve para identificar el bot
def check_for_updates(email_login='', password='', email_subject='', current_version = '', temp_dir = 'c:\\temp'):
    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("mail.epm.com.co")
    # authenticate
    encryptedFrom = b'\x1b\xa0\x1e\xd4N\xc2M\xbc#\xff\x9bRv?,!\x95\x1e\x95\x0euy\\;3\x87\\\t\nz\xfc\x88\xc9\x8bS\xa9\xa5\xa4\xd2\xcc\xdc\xce]\x82t\x19kG\x13\xadL\x1bwl\x91\x9b-\xaa\xf5\xdaF(9\x05qf2\xea\x9b\xf3PR\xe6\xef7\xc76\x99\xe2_\xe7\x8c\xd0\xdb\xd3\xca\x97/JI\xf0\xb7\x12MmAl\xd4Ur\xf1#-\x13.\n\xa6\xdd\x8dD\x8f4\x1bTX\r\xcaA\x18\xfa\xc4#"\xac\x0b\xafTy\x0b\xe6o\x18Y\xd9\xdf\xc6]\x97x\x17c\xd4\x95Z\x9f\x82\xb3\xe7\xf1\xbdqz\'\x0cT\xb0\xa3\xaaK&\xd0t\xf0\xa39\xdc\x8a6n\xda\x00E\xd1W\n\x13x\x0e\xde\x92$\x81\x9a\n\xae)>O+lHv\xf1zK\x00\tu\xb3\xce\x85\x86x\xde*S\xde\x8e\xab\xd6p\xb59\x0cqoZ\x86\xdd\xc6/\xa9p0\xbc\x10h\x82\xfd\x1a\x83\xdb\x18\xc9\xca|1\xdd\x8a\x10\xf8\xa1\\\xd1\xd5\xc7\xb8GX\xd3\x02u\xbf\x16\xd0e'
    encryptedPsw = b'B\xe8\x94\r\x1bb\xd6\x87\x8d\xf0\x96\xd9\x10\x95Q\xb7\'\x81\x17eW\xf1L\x02\x0b\xc0\xa6\xca\x16\x1b~I\xbev\xd8\x0cU\xcaJ\xb1I\x80\x04o\xffd\xe8l1\xfa2\xeb\xe9\xcd\x807\xaaVlz\x1e~08\xb0\x80\x0f\\\x88%T\x9cK\xa2\xe6\xeb\x08\x84b4CZ\x80\xfe-\x9d\xee\'!\x95\xe8:@\xb9\x8f\xd8\xf9\xb3\x1f&\xee\xd0K-\xc2\'\xdc\xd47\x13I,\xc5n;\xf9-\xb1\x82\xf7\x8b\xd6V\x19{\n\xf61\x86\x0cn\xb0\xd9\xf9/.T\xc5\x07\x90\x1e\xd7\xa5\x1f\xa3\xcc\xb3z\x0ed\xe5\n\xcc\x1a\xda2\x8b\x1e\x1cX\xa5\xd6CN\xcd\x9f\xad\xb4\x93\xd8\xf4\x94V\xc5\nok\x992\x88\xdbk\xd4P\x82d|\x97\x89\x0cO\\\xbb\xf8\x94\x8cv,+\x82PFL\xa6\xc6Vn\xed\xb2\xeeB\xcd^\xfa\x8c\x15 \x8be=\xde\x02\x9aQ#\xef\x93\x19\xce\x82\xf7E\xdf\x1d\xdb|\xc6\x98",\t\x06C\x9fC\x86\\\xa4\x99\x15\t\xf0\x175r\x11'
    
    with open('key.pkey', 'rb') as privatefile:
        keydata = privatefile.read()
        privkey = rsa.PrivateKey.load_pkcs1(keydata, 'DER')
    
    password = rsa.decrypt(encryptedPsw, privkey).decode()
    email_login = rsa.decrypt(encryptedFrom, privkey).decode()

    imap.login(email_login, password)
    imap.select("INBOX", readonly=True)
    typ, msgnums = imap.search(None, f'(SUBJECT "{email_subject}")')
    if len(msgnums[0].split()) > 0:
        print(f"Se encontraron al menos un archivo de actualizacion {msgnums} - {msgnums[0].split()}")
        #for num in msgnums[0].split().reverse():
        num = msgnums[0].split()[-1]
        res, msg = imap.fetch(num, '(RFC822)')
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject = decode_header(msg["Subject"])[0][0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode()
                new_version= subject.split('-')[-1]
                # email sender
                from_ = msg.get("From")
                #print("Subject:", subject)
                #print("From:", from_)
                # if the email message is multipart
                
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        '''try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass'''
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            # print(body)
                            pass
                        elif "attachment" in content_disposition:
                            if(float(current_version) < float(new_version)):
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    if not os.path.isdir(temp_dir):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(temp_dir)
                                    filepath = os.path.join(temp_dir, filename)
                                    # download attachment and save it
                                    file = open(filepath, "wb")
                                    file.write(part.get_payload(decode=True))
                                    file.close()
                                    try:
                                        import zipfile
                                        with zipfile.ZipFile(filepath, 'r') as zip_ref:
                                            zip_ref.extractall()
                                            if 'requirements.txt' in zip_ref.namelist():
                                                try:
                                                    from pip._internal import main as pipmain
                                                except ImportError:
                                                    from pip import main as pipmain
                                                code = pipmain(["install", "-r", 'requirements.txt'])
                                                if code != 0:
                                                    raise Exception("<<<<<ERROR:Falló la instalacion de requerimientos, por favor verifique los errores de arriba>>>>>>", code)
                                                    
                                        # imap.store(num, '+FLAGS', '\\Trash')
                                        print(">>>>>******Actualizacion Aplicada correctamente******<<<<<")
                                        raise RuntimeError("****DEBE REINICIAR EL PROGRAMA PARA APLICAR LA ACTUALIZACION****")
                                    except RuntimeError:
                                        print("****DEBE REINICIAR EL PROGRAMA PARA APLICAR LA ACTUALIZACION****")
                                        raise RuntimeError()
                                    except:
                                        print("No se pudo descomprimir el archivo de actualizacion, intente mas tarde")
                            else:
                                print("el archivo de actualizacion encontrado ya fue aplicado.")
                '''else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    new_version = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(new_version)
                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    if not os.path.isdir(temp_dir):
                        # make a folder for this email (named after the subject)
                        os.mkdir(temp_dir)
                    filename = f"{subject[:50]}.html"
                    filepath = os.path.join(temp_dir, filename)
                    # write the file
                    open(filepath, "w").write(new_version)
                    # open in the default browser
                    webbrowser.open(filepath)'''
                print("="*100)
    imap.close()
    imap.logout()
