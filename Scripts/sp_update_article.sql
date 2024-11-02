UPDATE Articles
SET
	Name = @Name,
	Details = @Details,
	UpdateAt = @UpdateAt,
	CategoryId = @CategoryId
WHERE
	Id = @Id