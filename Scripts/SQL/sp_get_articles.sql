SELECT
	a.Id, a.Name, a.Details, a.CreateAt, a.UpdateAt, c.Name AS CategoryName, a.CategoryId
FROM
	Articles a
LEFT JOIN
	Categories c
ON
	a.CategoryId = c.Id