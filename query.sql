Select * from service_api;

-- END --

Select * from atm;

-- END --

Update users
set FULL_NAME='Ngô Duy Hiếu'
Where USER_NAME='ADMIN';

-- END --

Select a.ATM_ID, a.SERIAL_NUMBER, b.DESCRIPTION, b.ADDRESS, b.MANAGER_NAME, b.MANAGER_PHONE
From atm a join organization b on a.ORGANIZATION_ID=b.ORGANIZATION_ID;

