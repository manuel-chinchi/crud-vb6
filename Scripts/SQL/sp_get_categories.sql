SELECT 
    c.Id, c.Name, c.State, c.CreateAt, c.UpdateAt, count(a.id) as ArticlesCount
FROM 
    Categories c
LEFT JOIN Articles a on
    c.Id = a.CategoryId 
GROUP by 
    c.Id, c.Name, c.State, c.CreateAt, c.UpdateAt
ORDER by 
    c.Id ASC 
