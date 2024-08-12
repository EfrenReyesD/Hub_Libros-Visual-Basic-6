
USE hublibros;

-- Crear la tabla Users
CREATE TABLE Users (
    UserId INT PRIMARY KEY IDENTITY(1,1),
    Username VARCHAR(50) NOT NULL,
    Password VARCHAR(255) NOT NULL,
    FirstName VARCHAR(50) NOT NULL,
    LastName VARCHAR(50) NOT NULL,
    ProfilePictureUrl VARCHAR(255) NOT NULL DEFAULT '',
    CreatedAt DATETIME DEFAULT GETDATE()
);

-- Crear la tabla Books
CREATE TABLE Books (
    BookId INT PRIMARY KEY IDENTITY(1,1),
    Title VARCHAR(255) NOT NULL,
    Author VARCHAR(255) NOT NULL,
    Genre VARCHAR(100),
    PdfUrl VARCHAR(255),
    CoverImage VARCHAR(255),
    Description TEXT
);

-- Crear la tabla UserBooks
CREATE TABLE UserBooks (
    UserBookId INT PRIMARY KEY IDENTITY(1,1),
    UserId INT NOT NULL,
    BookId INT NOT NULL,
    IsRead BIT NOT NULL DEFAULT 0,
    IsFavorite BIT NOT NULL DEFAULT 0,
    IsDisliked BIT NOT NULL DEFAULT 0,
    CONSTRAINT FK_UserBooks_Users FOREIGN KEY (UserId) REFERENCES Users(UserId),
    CONSTRAINT FK_UserBooks_Books FOREIGN KEY (BookId) REFERENCES Books(BookId)
);
