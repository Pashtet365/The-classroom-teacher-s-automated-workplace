-- �������� ������� ��������
CREATE TABLE [dbo].[��������] (
    [id_���������]           INT           IDENTITY (1, 1) NOT NULL,
    [���]                    NVARCHAR (50) NOT NULL,
    [�������]                NVARCHAR (50) NOT NULL,
    [����_��������]          DATE          NOT NULL,
    [���_�������_�������]    NVARCHAR (50) NOT NULL,
    [������������_���������] NVARCHAR (50) NOT NULL,
    [���������_�_�����]      NVARCHAR (50) NOT NULL,
    [���������_��_�����]     NVARCHAR (50) NOT NULL,
    [���������_�������]      NVARCHAR (50) NOT NULL,
    [������������]           NVARCHAR(50)  NOT NULL, 
    [������_��������]        NVARCHAR(50)  NOT NULL, 
    [����]                   NVARCHAR(50)  NOT NULL, 
    [������]                 NVARCHAR(50)  NOT NULL, 
    [������]                 NVARCHAR(50)  NOT NULL, 
    PRIMARY KEY CLUSTERED ([id_���������] ASC)
);

-- �������� ������� �����
CREATE TABLE [dbo].[�����] (
    [id_������]                  INT           IDENTITY (1, 1) NOT NULL,
    [������������]               NVARCHAR (50) NOT NULL,
    [����_����������_����������] DATE          NULL,
    PRIMARY KEY CLUSTERED ([id_������] ASC)
);

-- �������� ������� ��������
CREATE TABLE [dbo].[��������] (
    [id_��������]   INT           IDENTITY (1, 1) NOT NULL,
    [id_���������]  INT           NOT NULL,
    [���]           NVARCHAR (50) NOT NULL,
    [���]           NVARCHAR (50) NOT NULL,
    [����_��������] DATE          NOT NULL,
    [�����_������]  NVARCHAR(50)  NOT NULL, 
    [���������]     NVARCHAR(50)  NOT NULL, 
    [������]        NVARCHAR(50)  NOT NULL, 
    PRIMARY KEY CLUSTERED ([id_��������] ASC),
    CONSTRAINT [FK_��������_��������] FOREIGN KEY ([id_���������]) REFERENCES [dbo].[��������] ([id_���������])
);

-- �������� ������� ��������
CREATE TABLE [dbo].[��������] (
    [id_��������] INT           IDENTITY (1, 1) NOT NULL,
    [��������]    NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_��������] ASC)
);

-- �������� ������� �������
CREATE TABLE [dbo].[�������] (
    [id_�������]   INT  IDENTITY (1, 1) NOT NULL,
    [id_���������] INT  NOT NULL,
    [id_��������]  INT  NOT NULL,
    [�������]      INT  NOT NULL,
    [����]         DATE NOT NULL,
    PRIMARY KEY CLUSTERED ([id_�������] ASC),
    CONSTRAINT [FK_�������_��������] FOREIGN KEY ([id_���������]) REFERENCES [dbo].[��������] ([id_���������]),
    CONSTRAINT [FK_�������_��������] FOREIGN KEY ([id_��������]) REFERENCES [dbo].[��������] ([id_��������])
);

-- �������� ������� ��������
CREATE TABLE [dbo].[��������] (
    [id_��������]      INT  IDENTITY (1, 1) NOT NULL,
    [id_���������]     INT  NOT NULL,
    [����������_�����] INT  NOT NULL,
    [����]             DATE NOT NULL,
    PRIMARY KEY CLUSTERED ([id_��������] ASC),
    CONSTRAINT [FK_��������_��������] FOREIGN KEY ([id_���������]) REFERENCES [dbo].[��������] ([id_���������])
);

-- �������� ������� ��������
CREATE TABLE [dbo].[��������] (
    [id_��������] INT           IDENTITY (1, 1) NOT NULL,
    [id_��������] INT           NOT NULL,
    [����]        DATE          NOT NULL,
    [����]        NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_��������] ASC),
    CONSTRAINT [FK_��������_��������] FOREIGN KEY ([id_��������]) REFERENCES [dbo].[��������] ([id_��������])
);

-- �������� ������� �������
CREATE TABLE [dbo].[�������] (
    [id_�������]   INT           IDENTITY (1, 1) NOT NULL,
    [id_���������] INT           NOT NULL,
    [����_������]  DATE          NOT NULL,
    [����_�����]   DATE          NOT NULL,
    [���_�������]  NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([id_�������] ASC),
    CONSTRAINT [FK_�������_��������] FOREIGN KEY ([id_���������]) REFERENCES [dbo].[��������] ([id_���������])
);


-- ������� ������ � ������� ��������
INSERT INTO [dbo].[��������] ([���], [�������], [����_��������], [���_�������_�������], [������������_���������], [���������_�_�����], [���������_��_�����], [���������_�������], [������������], [������_��������], [����], [������], [������])
VALUES 
(N'������ ���� ��������', N'+79990000000', '2005-01-01', N'�������', N'�������', N'���', N'���', N'���', N'���', N'������', N'���', N'������', N'���'),
(N'������ ���� ��������', N'+79990000001', '2006-02-02', N'�������', N'�������', N'��', N'��', N'��', N'���', N'������', N'���', N'�������', N'��');

-- ������� ������ � ������� �����
INSERT INTO [dbo].[�����] ([������������], [����_����������_����������])
VALUES 
(N'1�', '2024-01-01'),
(N'2�', '2024-02-02');

-- ������� ������ � ������� ��������
INSERT INTO [dbo].[��������] ([id_���������], [���], [���], [����_��������], [�����_������], [���������], [������])
VALUES 
(1, N'������� ���� ��������', N'�', '1980-01-01', N'�������� �', N'��������', N'��. ������, �.1'),
(2, N'������� ����� ��������', N'�', '1982-02-02', N'�������� �', N'���������', N'��. �������, �.2');

-- ������� ������ � ������� ��������
INSERT INTO [dbo].[��������] ([��������])
VALUES 
(N'����������'),
(N'������� ����');

-- ������� ������ � ������� �������
INSERT INTO [dbo].[�������] ([id_���������], [id_��������], [�������], [����])
VALUES 
(1, 1, 5, '2024-01-15'),
(2, 2, 4, '2024-01-16');

-- ������� ������ � ������� ��������
INSERT INTO [dbo].[��������] ([id_���������], [����������_�����], [����])
VALUES 
(1, 2, '2024-01-10'),
(2, 3, '2024-01-11');

-- ������� ������ � ������� ��������
INSERT INTO [dbo].[��������] ([id_��������], [����], [����])
VALUES 
(1, '2024-01-20', N'������������ �������� 1'),
(2, '2024-01-21', N'������������ �������� 2');

-- ������� ������ � ������� �������
INSERT INTO [dbo].[�������] ([id_���������], [����_������], [����_�����], [���_�������])
VALUES 
(1, '2024-01-05', '2024-01-10', N'����������� �������'),
(2, '2024-01-06', '2024-01-11', N'������� �� �������');
