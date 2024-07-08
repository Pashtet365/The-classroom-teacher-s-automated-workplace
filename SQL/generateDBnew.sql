-- Создание таблицы Учащийся
CREATE TABLE [dbo].[Учащийся] (
    [id_учащегося]           INT           IDENTITY (1, 1) NOT NULL,
    [фио]                    NVARCHAR (50) NOT NULL,
    [телефон]                NVARCHAR (50) NOT NULL,
    [дата_рождения]          DATE          NOT NULL,
    [соц_бытовые_условия]    NVARCHAR (50) NOT NULL,
    [материальное_состояние] NVARCHAR (50) NOT NULL,
    [трудности_в_учёбе]      NVARCHAR (50) NOT NULL,
    [состояние_на_учёте]     NVARCHAR (50) NOT NULL,
    [нарушение_общения]      NVARCHAR (50) NOT NULL,
    [инвалидность]           NVARCHAR(50)  NOT NULL, 
    [группа_здоровья]        NVARCHAR(50)  NOT NULL, 
    [чаэс]                   NVARCHAR(50)  NOT NULL, 
    [кружки]                 NVARCHAR(50)  NOT NULL, 
    [сирота]                 NVARCHAR(50)  NOT NULL, 
    PRIMARY KEY CLUSTERED ([id_учащегося] ASC)
);

-- Создание таблицы Класс
CREATE TABLE [dbo].[Класс] (
    [id_класса]                  INT           IDENTITY (1, 1) NOT NULL,
    [наименование]               NVARCHAR (50) NOT NULL,
    [дата_последнего_обновления] DATE          NULL,
    PRIMARY KEY CLUSTERED ([id_класса] ASC)
);

-- Создание таблицы Родители
CREATE TABLE [dbo].[Родители] (
    [id_родителя]   INT           IDENTITY (1, 1) NOT NULL,
    [id_учащегося]  INT           NOT NULL,
    [фио]           NVARCHAR (50) NOT NULL,
    [пол]           NVARCHAR (50) NOT NULL,
    [дата_рождения] DATE          NOT NULL,
    [место_работы]  NVARCHAR(50)  NOT NULL, 
    [должность]     NVARCHAR(50)  NOT NULL, 
    [адресс]        NVARCHAR(50)  NOT NULL, 
    PRIMARY KEY CLUSTERED ([id_родителя] ASC),
    CONSTRAINT [FK_Родители_Учащийся] FOREIGN KEY ([id_учащегося]) REFERENCES [dbo].[Учащийся] ([id_учащегося])
);

-- Создание таблицы Предметы
CREATE TABLE [dbo].[Предметы] (
    [id_предмета] INT           IDENTITY (1, 1) NOT NULL,
    [название]    NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_предмета] ASC)
);

-- Создание таблицы Отметки
CREATE TABLE [dbo].[Отметки] (
    [id_отметки]   INT  IDENTITY (1, 1) NOT NULL,
    [id_учащегося] INT  NOT NULL,
    [id_предмета]  INT  NOT NULL,
    [отметка]      INT  NOT NULL,
    [дата]         DATE NOT NULL,
    PRIMARY KEY CLUSTERED ([id_отметки] ASC),
    CONSTRAINT [FK_Отметки_Учащийся] FOREIGN KEY ([id_учащегося]) REFERENCES [dbo].[Учащийся] ([id_учащегося]),
    CONSTRAINT [FK_Отметки_Предметы] FOREIGN KEY ([id_предмета]) REFERENCES [dbo].[Предметы] ([id_предмета])
);

-- Создание таблицы Пропуски
CREATE TABLE [dbo].[Пропуски] (
    [id_пропуска]      INT  IDENTITY (1, 1) NOT NULL,
    [id_учащегося]     INT  NOT NULL,
    [количество_часов] INT  NOT NULL,
    [дата]             DATE NOT NULL,
    PRIMARY KEY CLUSTERED ([id_пропуска] ASC),
    CONSTRAINT [FK_Пропуски_Учащийся] FOREIGN KEY ([id_учащегося]) REFERENCES [dbo].[Учащийся] ([id_учащегося])
);

-- Создание таблицы Собрание
CREATE TABLE [dbo].[Собрание] (
    [id_собрания] INT           IDENTITY (1, 1) NOT NULL,
    [id_родителя] INT           NOT NULL,
    [дата]        DATE          NOT NULL,
    [тема]        NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_собрания] ASC),
    CONSTRAINT [FK_Собрание_Родители] FOREIGN KEY ([id_родителя]) REFERENCES [dbo].[Родители] ([id_родителя])
);

-- Создание таблицы Справка
CREATE TABLE [dbo].[Справка] (
    [id_справки]   INT           IDENTITY (1, 1) NOT NULL,
    [id_учащегося] INT           NOT NULL,
    [дата_начала]  DATE          NOT NULL,
    [дата_конца]   DATE          NOT NULL,
    [вид_справки]  NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_справки] ASC),
    CONSTRAINT [FK_Справка_Учащийся] FOREIGN KEY ([id_учащегося]) REFERENCES [dbo].[Учащийся] ([id_учащегося])
);


-- Вставка данных в таблицу Учащийся
INSERT INTO [dbo].[Учащийся] ([фио], [телефон], [дата_рождения], [соц_бытовые_условия], [материальное_состояние], [трудности_в_учёбе], [состояние_на_учёте], [нарушение_общения], [инвалидность], [группа_здоровья], [чаэс], [кружки], [сирота])
VALUES 
(N'Иванов Иван Иванович', N'+79990000000', '2005-01-01', N'Хорошие', N'Средний', N'Нет', N'Нет', N'Нет', N'Нет', N'Первая', N'Нет', N'Футбол', N'Нет'),
(N'Петров Петр Петрович', N'+79990000001', '2006-02-02', N'Средние', N'Хороший', N'Да', N'Да', N'Да', N'Нет', N'Вторая', N'Нет', N'Шахматы', N'Да');

-- Вставка данных в таблицу Класс
INSERT INTO [dbo].[Класс] ([наименование], [дата_последнего_обновления])
VALUES 
(N'1А', '2024-01-01'),
(N'2Б', '2024-02-02');

-- Вставка данных в таблицу Родители
INSERT INTO [dbo].[Родители] ([id_учащегося], [фио], [пол], [дата_рождения], [место_работы], [должность], [адресс])
VALUES 
(1, N'Иванова Анна Петровна', N'Ж', '1980-01-01', N'Компания А', N'Менеджер', N'ул. Ленина, д.1'),
(2, N'Петрова Мария Ивановна', N'Ж', '1982-02-02', N'Компания Б', N'Бухгалтер', N'ул. Пушкина, д.2');

-- Вставка данных в таблицу Предметы
INSERT INTO [dbo].[Предметы] ([название])
VALUES 
(N'Математика'),
(N'Русский язык');

-- Вставка данных в таблицу Отметки
INSERT INTO [dbo].[Отметки] ([id_учащегося], [id_предмета], [отметка], [дата])
VALUES 
(1, 1, 5, '2024-01-15'),
(2, 2, 4, '2024-01-16');

-- Вставка данных в таблицу Пропуски
INSERT INTO [dbo].[Пропуски] ([id_учащегося], [количество_часов], [дата])
VALUES 
(1, 2, '2024-01-10'),
(2, 3, '2024-01-11');

-- Вставка данных в таблицу Собрание
INSERT INTO [dbo].[Собрание] ([id_родителя], [дата], [тема])
VALUES 
(1, '2024-01-20', N'Родительское собрание 1'),
(2, '2024-01-21', N'Родительское собрание 2');

-- Вставка данных в таблицу Справка
INSERT INTO [dbo].[Справка] ([id_учащегося], [дата_начала], [дата_конца], [вид_справки])
VALUES 
(1, '2024-01-05', '2024-01-10', N'Медицинская справка'),
(2, '2024-01-06', '2024-01-11', N'Справка по болезни');
