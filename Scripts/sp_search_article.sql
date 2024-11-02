SELECT a.Id, a.Name, a.Details, a.CreateAt, a.UpdateAt, c.Name AS CategoryName, (a.id || a.Name || a.Details || c.Name) as ROWSTRING
FROM Articles AS a, Categories AS c
WHERE 
	a.CategoryId = c.Id AND
	ROWSTRING like '%' || @Search || '%'