--- test data for database
INSERT INTO Categories 
			(Name, State, CreateAt, UpdateAt) 
	VALUES 	
			('Otro', 1, datetime('now'), NULL),
			('Ropa Masculina', 1, datetime('now'), NULL),
			('Ropa Femenina', 1, datetime('now'), NULL),
			('Ropa Unisex', 1, datetime('now'), NULL),
			('Calzado', 1, datetime('now'), NULL)

INSERT INTO Articles
			(Name, Details, CreateAt, UpdateAt, IdCategory)
	VALUES
			('Buzo t/canguro', '5xU', datetime('now'), NULL, 2),
			('Jean elastizado', '10xU', datetime('now'), NULL, 2),
			('Gorra blanca', '15xU', datetime('now'), NULL, 2)

			
			
SELECT * from Categories 
select * from Articles 			
DELETE from 'Categories'  ---borrar todos registros

drop TABLE Categories
drop TABLE Articles

strftime('%m/%d/%Y', date('now')) 
date('now')