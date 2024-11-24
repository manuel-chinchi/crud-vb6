-- database schema for proyect crud_vb6. source: https://github.com/manuel-chinchi/crud-vb6
CREATE TABLE Categories (
	Id INTEGER PRIMARY KEY,
	Name TEXT NOT NULL,
	State INTEGER DEFAULT 1, --true=1|false=0
	CreateAt TEXT,
	UpdateAt TEXT 
)

CREATE TABLE Articles (
	Id INTEGER PRIMARY KEY,
	Name TEXT NOT NULL,
	Details TEXT,
	CreateAt TEXT,
	UpdateAt TEXT,
	CategoryId INTEGER,
	FOREIGN KEY (CategoryId) REFERENCES Categories(Id) ON DELETE RESTRICT
)