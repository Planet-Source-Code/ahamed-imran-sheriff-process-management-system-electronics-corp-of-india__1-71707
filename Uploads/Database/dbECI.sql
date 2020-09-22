

---------------------------------RDSD Assignment--------------------------------


--Here I am creating a database and establishing the logical filename 
--that SQL Server will use to reference this database. I am also setting the file name
--that the operating system will use. Then, I have established the starting size in MB 
--and the rate of growth of the database when the server requests more room.
CREATE DATABASE [dbECI]
ON (NAME = N'dbECI_Data'
,FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\Data\dbECI_Data.MDF' 
,SIZE = 2, FILEGROWTH = 10%) 
LOG ON (NAME = N'dbECI_Log'
 ,FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\Data\dbECI_Log.LDF'
,SIZE = 2
,FILEGROWTH = 10%)   


--Using the database.
use dbECI


----- Create Table Statements -----

create table UserAccount(
	UserID char(5) check(UserID between 'OP001' and 'OP999') primary key,
	UserName varchar(50) check (UserName not like '%[0-9]%') not null,	
	Password varchar(50) not null,
	AccessLevel varchar(50) check (AccessLevel not like '%[0-9]%') not null,	
)

create table ItemDetails(
	Item_ID char(5) check(Item_ID between 'II001' and 'II999') primary key,
	Item_Type varchar(50) check(Item_Type in('Television', 'Washing Machine', 'Refrigerator', 'Video Record Player')) not null,	
	Item_Size varchar(10) check(Item_Size in('Small', 'Medium', 'Large')) not null,
	Manufacturer varchar(50) check (Manufacturer not like '%[0-9]%') not null,	
	Cost_Price  float check (Cost_Price  > 0 AND Cost_Price  < 999999) not null,
	Retail_Price float check (Retail_Price  > 0 AND Retail_Price  < 999999) not null,
	Stock_Count int not null,	
)

create table CustomerDetails(
	Customer_ID char(5) check(Customer_ID between 'CD001' and 'CD999') primary key,
	Customer_First_Name varchar(30) check (Customer_First_Name not like '%[0-9]%') not null,	
	Customer_Last_Name varchar(30) check (Customer_Last_Name not like '%[0-9]%') not null,	
	Address varchar(100) check(len(Address)>0)not null,
	Telephone_Number varchar(15) check (Telephone_Number not like '%[a-z]%') 	
)

create table WarehouseDetails(
	Warehouse_Location varchar(30) primary key,
	Address varchar(100) check(len(Address)>0) not null,
	Telephone_Number varchar(15) check (Telephone_Number not like '%[a-z]%') 
)

create table ServiceDepot(
	Service_Depot_ID char(5) check(Service_Depot_ID between 'SD001' and 'SD999') primary key,
	Service_Depot_Location varchar(30) check (Service_Depot_Location not like '%[0-9]%') not null,	
	Address varchar(100) check(len(Address)>0) not null,
)

create table Engineer(
	Engineer_ID char(5) check(Engineer_ID between 'EN001' and 'EN999') primary key,
	Engineer_Name varchar(60) check (Engineer_Name not like '%[0-9]%') not null,	
	Address varchar(100) check(len(Address)>0) not null,
	Telephone_Number varchar(15) check (Telephone_Number not like '%[a-z]%'), 
	Service_Depot_ID char(5) not null,
	foreign key(Service_Depot_ID) references ServiceDepot,	
)

create table RetailStoreDetails(
	Retail_Store_ID char(5) check(Retail_Store_ID between 'RS001' and 'RS999') primary key,
	Retail_Store_Name varchar(30) check (Retail_Store_Name not like '%[0-9]%') not null,
	Retail_Store_Location varchar(30) check (Retail_Store_Location not like '%[0-9]%') not null,
	Address varchar(100) check(len(Address)>0) not null,
	Telephone_Number varchar(15) check (Telephone_Number not like '%[a-z]%'), 
	Service_Depot_ID char(5) not null,
	Warehouse_Location varchar(30) check (Warehouse_Location not like '%[0-9]%') not null,	
	foreign key(Service_Depot_ID) references ServiceDepot,	
	foreign key(Warehouse_Location) references WarehouseDetails,	
)

create table PurchaseDetails(
	Purchase_ID char(5) check(Purchase_ID between 'PD001' and 'PD999') primary key,
	Purchase_Date datetime check (Purchase_Date >= getDate()-1) not null,
	Retail_Store_ID char(5) not null,	
	Customer_ID char(5) not null,
	Item_ID char(5) not null,	
	Quantity smallint default 1 not null,
	Total_Cost float check (Total_Cost  > 0 AND Total_Cost  < 999999) not null,
	foreign key(Item_ID) references ItemDetails,	
	foreign key(Customer_ID) references CustomerDetails on update cascade,	
	foreign key(Retail_Store_ID) references RetailStoreDetails,	
)

create table ServiceAgreementDetails(
	Service_ID char(5) check(Service_ID between 'SA001' and 'SA999') primary key,
	Service_Depot_ID char(5) not null,
	Purchase_ID char(5) not null,
	Customer_ID char(5) not null,
	Purchase_Date datetime check (Purchase_Date >= getDate()-1) not null,
	Duration smallint default 3 not null,
	Quantity smallint default 1 not null,
	Total_Cost float check (Total_Cost  > 0 AND Total_Cost  < 999999) not null,
	foreign key(Purchase_ID) references PurchaseDetails,
	foreign key(Customer_ID) references CustomerDetails,		
	foreign key(Service_Depot_ID) references ServiceDepot,	
)

create table MaintenanceRecord(
	Maintenance_ID char(5) check(Maintenance_ID between 'MR001' and 'MR999') primary key,
	Service_ID char(5) not null,
	Service_Depot_ID char(5) not null,
	Engineer_ID char(5) not null,	
	Service_Date datetime check (Service_Date >= getDate()-1) not null,
	Maintenance_Cost float check (Maintenance_Cost  > 0 AND Maintenance_Cost  < 999999) not null,
	foreign key(Service_ID) references ServiceAgreementDetails,
	foreign key(Service_Depot_ID) references ServiceDepot,
	foreign key(Engineer_ID) references Engineer,	
)




------------User Stored Procesdures----------

CREATE PROCEDURE usp_ItemSearch as
SELECT Item_ID,Item_Type,Item_Size,Manufacturer,Retail_Price
FROM ItemDetails
ORDER BY Item_ID

CREATE PROCEDURE usp_CustomerSearch as
SELECT Customer_ID,Customer_First_Name,Customer_Last_Name
FROM CustomerDetails
ORDER BY Customer_ID


--------------------Views----------------------
CREATE VIEW PurchasesView
AS
SELECT   Purchase_ID,Purchase_Date,Customer_ID,Item_ID
FROM     PurchaseDetails


CREATE VIEW RetailStoreView
AS
SELECT   Retail_Store_ID,Retail_Store_Name,Retail_Store_Location
FROM     RetailStoreDetails


-----------Purchases Insert Trigger-----------
CREATE TABLE [dbo].[PurchasesTrigger] (
	Purchase_Date datetime not null,
	Total_Cost float not null,
)

-- Start of command
CREATE TRIGGER tr_InsertPurchases 
-- The table name I want to affect
ON PurchaseDetails 
-- The type of trigger I want
FOR INSERT 
AS
-- I'll set up two variables to hold the date
-- and cost
DECLARE @Purchase_Date datetime
DECLARE @Total_Cost float

-- Now I'll make use of the inserted virtual table I mentioned, 
-- setting the values of the variables to the data the user
-- sends. 
SELECT @Purchase_Date = (SELECT Purchase_Date FROM Inserted)
SELECT @Total_Cost = (SELECT Total_Cost FROM Inserted)
-- And now I'll use those variables to insert data into
-- the TriggerTest Table I made earlier 
INSERT PurchasesTrigger values (@Purchase_Date, @Total_Cost)


---- UserAccount Sample Data ----

insert into UserAccount(UserID,UserName,Password,AccessLevel)
values('OP001','Imran Sheriff','operator','Operator')

insert into UserAccount(UserID,UserName,Password,AccessLevel)
values('OP002','Dinithi Vithanage','operator','Operator')

insert into UserAccount(UserID,UserName,Password,AccessLevel)
values('OP003','Salvin Saleh','operator','Operator')

insert into UserAccount(UserID,UserName,Password,AccessLevel)
values('OP004','Imthiaz Sheriff','admin','Administrator')



----ItemDetails Sample Data ----

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II001','Television','Small','Samson Global',12000,16000,10)

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II002','Television','Medium','Samson Global',16000,20000,10)

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II003','Television','Large','Samson Global',20000,24000,10)

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II004','Video Record Player','Small','AudioX Electronics',2000,4000,10)

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II005','Video Record Player','Medium','AudioX Electronics',4000,6000,10)

insert into ItemDetails(Item_ID,Item_Type,Item_Size,Manufacturer,Cost_Price,Retail_Price,Stock_Count)
values('II006','Video Record Player','Large','AudioX Electronics',6000,8000,10)



----CustomerDetails Sample Data ----

insert into CustomerDetails(Customer_ID,Customer_First_Name,Customer_Last_Name,Address,Telephone_Number)
values('CD001','Fazna','Hudah','100, Shrubbery Gardens, Colombo - 04.','011 2625691')

insert into CustomerDetails(Customer_ID,Customer_First_Name,Customer_Last_Name,Address,Telephone_Number)
values('CD002','Vindu','Palihakkara','34, Dawson Street, Colombo - 03.','011 2696872')

insert into CustomerDetails(Customer_ID,Customer_First_Name,Customer_Last_Name,Address,Telephone_Number)
values('CD003','Sarah','Sheriff','55/2, Davidson Road, Colombo - 04.','011 2593673')

insert into CustomerDetails(Customer_ID,Customer_First_Name,Customer_Last_Name,Address,Telephone_Number)
values('CD004','Abheetha','Rathnayake','502/8, Baudhaloka Mawatha, Colombo - 06.','011 2596440')

insert into CustomerDetails(Customer_ID,Customer_First_Name,Customer_Last_Name,Address,Telephone_Number)
values('CD005','Thivanka','Makalanda','400/7, Rexton Avenue, Colombo - 2.','011 2523568')



----WarehouseDetails Sample Data----
insert into WarehouseDetails(Warehouse_Location,Address,Telephone_Number)
values ('Bambalapitiya','No 508/6, Golden Terrace','011 2659878')

insert into WarehouseDetails(Warehouse_Location,Address,Telephone_Number)
values ('Wellawatta','No 98, 1st Lane','011 2969205')

insert into WarehouseDetails(Warehouse_Location,Address,Telephone_Number)
values ('Jaffna','50, Shrewberry Avenue','011 2756489')

insert into WarehouseDetails(Warehouse_Location,Address,Telephone_Number)
values ('Kandy','No 4, Red Wood Place','011 2654784')

insert into WarehouseDetails(Warehouse_Location,Address,Telephone_Number)
values ('Anuradhapura','No 12, SandShell Rd','011 2487962')


---ServiceDepot Sample Data---

insert into ServiceDepot(Service_Depot_ID, Service_Depot_Location, Address)
values('SD001','Kandy','No 2 , Sherwood Loc')

insert into ServiceDepot(Service_Depot_ID, Service_Depot_Location, Address)
values('SD002','Anuradhapura','No 100 , Mile Avenue, Ration Road')

insert into ServiceDepot(Service_Depot_ID, Service_Depot_Location, Address)
values('SD003','Jaffna','No 58 , Shelwood Place')

insert into ServiceDepot(Service_Depot_ID, Service_Depot_Location, Address)
values('SD004','Bambalapitiya','No 81 , Railplace, Shelton Road')

insert into ServiceDepot(Service_Depot_ID, Service_Depot_Location, Address)
values('SD005','Wellawatta','No 56 , Leyton Terrace')



---Engineer Sample Data---
insert into Engineer (Engineer_ID,Engineer_Name,Address,Telephone_Number, Service_Depot_ID)
values ('EN001','Paul Perera','No 34, Jaded Terrace','011 2987456','SD001')

insert into Engineer (Engineer_ID,Engineer_Name,Address,Telephone_Number, Service_Depot_ID)
values ('EN002','Shania Twain','No 69, Chubby Road','011 2552005','SD005')

insert into Engineer (Engineer_ID,Engineer_Name,Address,Telephone_Number, Service_Depot_ID)
values ('EN003','James Lafferty','No 89/7, Wilson Place','011 2551551','SD003')

insert into Engineer (Engineer_ID,Engineer_Name,Address,Telephone_Number, Service_Depot_ID)
values ('EN004','Shazna Mowlana','No 34, Mystique Grove','011 2290023','SD002')

insert into Engineer (Engineer_ID,Engineer_Name,Address,Telephone_Number, Service_Depot_ID)
values ('EN005','Barack Obama','No 87, Pennsylvania Avenue','011 2444555','SD004')



---RetailStoreDetails Sample Data---
insert into RetailStoreDetails (Retail_Store_ID, Retail_Store_Name, Retail_Store_Location, Address, Telephone_Number, Service_Depot_ID, Warehouse_Location)
values('RS001','Shamil and Brothers','Wellawatta','No 09, Calcutta Avenue','011 2723895','SD005','Wellawatta')

insert into RetailStoreDetails (Retail_Store_ID, Retail_Store_Name, Retail_Store_Location, Address, Telephone_Number, Service_Depot_ID, Warehouse_Location)
values('RS002','Rushan Pvt Ltd','Kandy','No 55, ShineRay','011 2410710','SD001','Kandy')

insert into RetailStoreDetails (Retail_Store_ID, Retail_Store_Name, Retail_Store_Location, Address, Telephone_Number, Service_Depot_ID, Warehouse_Location)
values('RS003','Pigtopus Traders','Jaffna','No 08,Lewis Road','011 2794007','SD003','Jaffna')

insert into RetailStoreDetails (Retail_Store_ID, Retail_Store_Name, Retail_Store_Location, Address, Telephone_Number, Service_Depot_ID, Warehouse_Location)
values('RS004','Blue Ray Traders','Bambalapitiya','No 502/9,SweetLess Avenue','011 2248503','SD002','Bambalapitiya')

insert into RetailStoreDetails (Retail_Store_ID, Retail_Store_Name, Retail_Store_Location, Address, Telephone_Number, Service_Depot_ID, Warehouse_Location)
values('RS005','Mushtak Limited','Anuradhapura','No 65,Shimmering Avenue','011 2520051','SD004','Anuradhapura')



---PurchaseDetails Sample Data---

insert into PurchaseDetails (Purchase_ID,Purchase_Date,Retail_Store_ID,Customer_ID,Item_Id,Quantity,Total_Cost)
values('PD001','05/25/2009','RS002','CD005','II005',2,16000)

insert into PurchaseDetails (Purchase_ID,Purchase_Date,Retail_Store_ID,Customer_ID,Item_Id,Quantity,Total_Cost)
values('PD002','05/26/2009','RS004','CD003','II004',1,4000)

insert into PurchaseDetails (Purchase_ID,Purchase_Date,Retail_Store_ID,Customer_ID,Item_Id,Quantity,Total_Cost)
values('PD003','05/28/2009','RS002','CD005','II006',2,32000)

insert into PurchaseDetails (Purchase_ID,Purchase_Date,Retail_Store_ID,Customer_ID,Item_Id,Quantity,Total_Cost)
values('PD004','06/01/2009','RS001','CD004','II003',1,24000)

insert into PurchaseDetails (Purchase_ID,Purchase_Date,Retail_Store_ID,Customer_ID,Item_Id,Quantity,Total_Cost)
values('PD005','06/04/2009','RS003','CD002','II002',4,80000)



---ServiceAgreementDetails Sample Data---
insert into ServiceAgreementDetails (Service_ID,Service_Depot_ID,Purchase_ID,Customer_ID,Purchase_Date,Duration,Quantity,Total_Cost)
values('SA001','SD001','PD001','CD001','05/25/2009',3,2,6000)

insert into ServiceAgreementDetails (Service_ID,Service_Depot_ID,Purchase_ID,Customer_ID,Purchase_Date,Duration,Quantity,Total_Cost)
values('SA002','SD002','PD002','CD002','05/26/2009',3,1,3000)

insert into ServiceAgreementDetails (Service_ID,Service_Depot_ID,Purchase_ID,Customer_ID,Purchase_Date,Duration,Quantity,Total_Cost)
values('SA003','SD002','PD003','CD003','05/28/2009',3,2,6000)

insert into ServiceAgreementDetails (Service_ID,Service_Depot_ID,Purchase_ID,Customer_ID,Purchase_Date,Duration,Quantity,Total_Cost)
values('SA004','SD003','PD004','CD004','06/01/2009',3,1,3000)

insert into ServiceAgreementDetails (Service_ID,Service_Depot_ID,Purchase_ID,Customer_ID,Purchase_Date,Duration,Quantity,Total_Cost)
values('SA005','SD005','PD005','CD005','06/04/2009',3,4,12000)



---MaintenanceRecord SampleData---

insert into MaintenanceRecord (Maintenance_ID,Service_ID,Service_Depot_ID,Engineer_ID,Service_Date,Maintenance_Cost)
values ('MR001','SA001','SD001','EN001','06/10/2009',2000)

insert into MaintenanceRecord (Maintenance_ID,Service_ID,Service_Depot_ID,Engineer_ID,Service_Date,Maintenance_Cost)
values ('MR002','SA002','SD001','EN001','07/10/2009',2000)

insert into MaintenanceRecord (Maintenance_ID,Service_ID,Service_Depot_ID,Engineer_ID,Service_Date,Maintenance_Cost)
values ('MR003','SA003','SD001','EN002','08/10/2009',2000)

insert into MaintenanceRecord (Maintenance_ID,Service_ID,Service_Depot_ID,Engineer_ID,Service_Date,Maintenance_Cost)
values ('MR004','SA004','SD001','EN003','09/10/2009',2000)

insert into MaintenanceRecord (Maintenance_ID,Service_ID,Service_Depot_ID,Engineer_ID,Service_Date,Maintenance_Cost)
values ('MR005','SA005','SD001','EN004','10/10/2009',2000)





