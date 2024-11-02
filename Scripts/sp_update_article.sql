update Articles
set
	Name=@Name,
	Details=@Details,
	UpdateAt=@UpdateAt,
	CategoryId=@CategoryId
where
	Id=@Id