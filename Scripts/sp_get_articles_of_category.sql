SELECT 
	Id, Name, Details, CreateAt, UpdateAt, CategoryId 
FROM Articles WHERE CategoryId = @CategoryId
