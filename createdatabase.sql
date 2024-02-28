-- Создание таблицы "Students"
CREATE TABLE Students (
    StudentID INT PRIMARY KEY IDENTITY(1,1),
    FirstName NVARCHAR(50),
    LastName NVARCHAR(50),
    Age INT,
    Phone NVARCHAR(20)
);

-- Создание таблицы "Locations"
CREATE TABLE Locations (
    LocationID INT PRIMARY KEY IDENTITY(1,1),
    Name NVARCHAR(100)
);

-- Создание таблицы "LessonTypes"
CREATE TABLE LessonTypes (
    LessonTypeID INT PRIMARY KEY IDENTITY(1,1),
    Name NVARCHAR(100)
);

-- Создание таблицы "Schedule"
CREATE TABLE Schedule (
    ScheduleID INT PRIMARY KEY IDENTITY(1,1),
    LessonDate DATE,
    StartTime TIME,
    EndTime TIME,
    StudentID INT FOREIGN KEY REFERENCES Students(StudentID),
    LocationID INT FOREIGN KEY REFERENCES Locations(LocationID),
    LessonTypeID INT FOREIGN KEY REFERENCES LessonTypes(LessonTypeID),
    Grade INT -- Оценка за занятие
);