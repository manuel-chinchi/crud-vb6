INSERT INTO Articles 
	(Name, Details, CreateAt, UpdateAt, CategoryId)
VALUES
	(@Name, @Details, @CreateAt, NULL, @CategoryId)