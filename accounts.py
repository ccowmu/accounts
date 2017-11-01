import ldap
import time
import smtplib
import getpass
from email.mime.multipart import MIMEMultipart

#TODO: exception handling, closing connections, more functionality?? cut into modules, PEP8, proper pythonic comments, proper password handling.

LDAP_URI = 'ldap://localhost:389'
MEMBER_BASE = 'cn=members,dc=yakko,dc=cs,dc=wmich,dc=edu'
POSIX_DAY = 86400 
EMAIL_PASSWORD = None
DESIRED_FIELDS = ['uid', 'shadowExpire', 'mail']

#_convert_day_to_shadow converts a unix timestamp plus day offset into a shadow timestamp and returns it as a string
def _convert_day_to_shadow(dayOffset):
    return str(int((time.time() + dayOffset * POSIX_DAY) / POSIX_DAY))


#_convert_shadow_to_unix converts a shadow timestamp into a unix timestamp
def _convert_shadow_to_unix(shadowTime):
    shadowTime = int(shadowTime)
    return str(int((shadowTime * POSIX_DAY)))


def _getpass_email_wrapper():
    global EMAIL_PASSWORD
    if EMAIL_PASSWORD is None:
        EMAIL_PASSWORD = getpass.getpass('Enter Email Password: ')


#get_expired_on_date queries ldap and returns all accounts w/ DESIRED_FIELDS that are set to expire on DATE
def get_expired_on_date(date):
    l = _connect_ldap()
    preparedQuery = '(shadowExpire=' + _convert_day_to_shadow(date) + ')'
    r = l.search_s(MEMBER_BASE, ldap.SCOPE_SUBTREE, preparedQuery, DESIRED_FIELDS) 
    del l

    if r:
    	return r


#get_expired_in_range returns a list of all expired ldap accounts between the specified [BEGINNING and END] range
def get_expired_in_range(beginning, end):
    if beginning < end:
        l = _connect_ldap()
    	preparedQuery = '(&(shadowExpire>=' + _convert_day_to_shadow(beginning) + ')' + '(shadowExpire<=' + _convert_day_to_shadow(end) + '))'
    	r = l.search_s(MEMBER_BASE, ldap.SCOPE_SUBTREE, preparedQuery, DESIRED_FIELDS)
        del l
    	return r


#is_expired takes a MEMBERUID and returns TRUE, FALSE, or 'Account does not exist'
def is_expired(memberUID):
    l = _connect_ldap()
    preparedQuery = '(uid=' + str(memberUID) + ')'
    r = l.search_s(MEMBER_BASE, ldap.SCOPE_SUBTREE, preparedQuery, ['shadowExpire'])
   
    del l 
    if not r:
        return 'Account does not exist'
   
    r = r[0][1].get('shadowExpire', '')

    if r[0] > _convert_day_to_shadow(0):
        return False
    else: 
	return True


#get_mail takes a MEMBERUID and returns the ldap mail entry for that user
def get_mail(memberUID):
    l = _connect_ldap()
    preparedQuery = '(uid=' + str(memberUID) + ')'
    r = l.search_s(MEMBER_BASE, ldap.SCOPE_SUBTREE, preparedQuery, ['mail'])

    del l
    if not r:
        return 'Email not on record or account does not exist'

    email = r[0][1].get('mail', '')
    return email


#get_expiration takes a MEMBERUID and returns their date of expiration as a unix timestamp
def get_expiration(memberUID):
    l = _connect_ldap()
    preparedQuery = '(uid=' + str(memberUID) + ')'
    r = l.search_s(MEMBER_BASE, ldap.SCOPE_SUBTREE, preparedQuery, ['shadowExpire'])
    
    del l
    if not r:
        return 'Email not on record or account does not exist'

    shadowExpire = r[0][1].get('shadowExpire', '')
    return _convert_shadow_to_unix(shadowExpire[0])


#email_expiry_notification takes an smtpConn object, memberEmail, and expiryDate and sends a formatted email to the designated address
def email_expiry_notification(memberEmail, expiryDate):

    s = _connect_mail()
    msg = MIMEMultipart()
    msg['Subject']='Account Expiration Approaching'
    msg['Body']='Your account is set to expire in' + expiryDate + 'days.'
    s.sendmail('rso_cclub@wmich.edu', memberEmail, msg.as_string())


#_connect_mail connects to the office365 mail server using cclub credentials and returns the connection object
def _connect_mail():
    s = smtplib.SMTP(host='smtp.office365.com', port=587)
    s.starttls()
    _getpass_email_wrapper()
    s.login('rso_cclub@wmich.edu', EMAIL_PASSWORD)
    return s

#_connect_ldap connects to the cclub ldap server and returns the connection object
def _connect_ldap():
    l = ldap.initialize(LDAP_URI)
    l.protocol_version = ldap.VERSION3
    l.simple_bind_s()
    return l
 
try:
#DEMO:
    keyDates = [30,15,10,5,2,1] 
    accounts = map(get_expired_on_date, keyDates) #use the keyDates list to gather accounts that are near expiration
    accounts = filter(lambda entry: entry, accounts) #remove nones from accounts list
    print "Printing accounts:\n\n"
    print accounts
    hoi = get_expired_in_range(0, 10)
    print hoi
    print is_expired('kami')
    print get_mail('flay')
    print get_expiration('flay')
except Exception, error:
    print error
