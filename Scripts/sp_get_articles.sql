SELECT
	a.Id, a.Name, a.Details, a.CreateAt, a.UpdateAt, c.Name AS CategoryName
FROM
	Articles AS a, Categories AS c
WHERE
	a.CategoryId = c.Id