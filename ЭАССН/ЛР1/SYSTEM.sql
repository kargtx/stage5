SELECT username, account_status, default_tablespace 
FROM dba_users 
WHERE username IN ('SYSTEM', 'HR', 'BANK');

CREATE USER bank IDENTIFIED BY bank
DEFAULT TABLESPACE TBS_DAN
TEMPORARY TABLESPACE temp
ACCOUNT UNLOCK;

GRANT DBA TO bank;




