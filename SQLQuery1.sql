-- =============================================
-- 1. إنشاء الجداول الأساسية (Database Schema)
-- =============================================

-- جدول العيادات (Clinics)
CREATE TABLE Clinics (
    ClinicCode VARCHAR(10) PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL,
    Address NVARCHAR(200),
    Phone VARCHAR(20)
);

-- جدول الكادر الطبي والموظفين (Staff)
CREATE TABLE Staff (
    EmployeeID VARCHAR(10) PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL,
    Role NVARCHAR(50) NOT NULL,
    YearsOfExperience INT,
    Password NVARCHAR(100) DEFAULT 'password123', -- كلمة المرور لتسجيل الدخول
    ClinicCode VARCHAR(10) FOREIGN KEY REFERENCES Clinics(ClinicCode)
);

-- جدول الأطباء (Doctors)
CREATE TABLE Doctors (
    EmployeeID VARCHAR(10) PRIMARY KEY FOREIGN KEY REFERENCES Staff(EmployeeID),
    MedicalSpecialty NVARCHAR(100),
    LicenseNumber VARCHAR(50)
);

-- جدول الممرضات (Nurses)
CREATE TABLE Nurses (
    EmployeeID VARCHAR(10) PRIMARY KEY FOREIGN KEY REFERENCES Staff(EmployeeID),
    CertificationLevel NVARCHAR(50)
);

-- جدول المرضى (Patients)
CREATE TABLE Patients (
    PatientID VARCHAR(20) PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL,
    DateOfBirth DATE
);

-- جدول أرقام هواتف المرضى (PatientPhones)
CREATE TABLE PatientPhones (
    PatientID VARCHAR(20) FOREIGN KEY REFERENCES Patients(PatientID),
    PhoneNumber VARCHAR(20),
    PRIMARY KEY (PatientID, PhoneNumber)
);

-- جدول الزيارات (Visits)
CREATE TABLE Visits (
    VisitID INT IDENTITY(1,1) PRIMARY KEY,
    PatientID VARCHAR(20) FOREIGN KEY REFERENCES Patients(PatientID),
    DoctorID VARCHAR(10) FOREIGN KEY REFERENCES Doctors(EmployeeID),
    ClinicCode VARCHAR(10) FOREIGN KEY REFERENCES Clinics(ClinicCode),
    VisitDate DATETIME DEFAULT GETDATE()
);

-- جدول التشخيصات (Diagnoses)
CREATE TABLE Diagnoses (
    DiagnosisCode VARCHAR(10) PRIMARY KEY,
    Description NVARCHAR(250) NOT NULL
);

-- جدول تشخيصات الزيارة (Visit_Diagnoses)
CREATE TABLE Visit_Diagnoses (
    VisitID INT FOREIGN KEY REFERENCES Visits(VisitID),
    DiagnosisCode VARCHAR(10) FOREIGN KEY REFERENCES Diagnoses(DiagnosisCode),
    PRIMARY KEY (VisitID, DiagnosisCode)
);

-- جدول الإحالات (Referrals)
CREATE TABLE Referrals (
    ReferralID INT IDENTITY(1,1) PRIMARY KEY,
    PatientID VARCHAR(20) FOREIGN KEY REFERENCES Patients(PatientID),
    FromDoctorID VARCHAR(10) FOREIGN KEY REFERENCES Doctors(EmployeeID),
    ToDoctorID VARCHAR(10) FOREIGN KEY REFERENCES Doctors(EmployeeID),
    ReferralDate DATETIME DEFAULT GETDATE(),
    Status NVARCHAR(20) DEFAULT 'Pending'
);

-- =============================================
-- 2. إدراج بيانات أولية وتجريبية (Initial Data)
-- =============================================

-- إضافة عيادة تجريبية
INSERT INTO Clinics (ClinicCode, Name, Address, Phone) 
VALUES ('CLN01', 'General Practice Clinic', 'Main Building - Floor 1', '065000000');

-- إضافة موظف/طبيب لتسجيل الدخول
INSERT INTO Staff (EmployeeID, Name, Role, YearsOfExperience, Password, ClinicCode) 
VALUES ('EMP01', 'Dr. Ahmed', 'Doctor', 8, 'password123', 'CLN01');

INSERT INTO Doctors (EmployeeID, MedicalSpecialty, LicenseNumber) 
VALUES ('EMP01', 'General Medicine', 'LIC-99823');

-- إضافة مريض تجريبي
INSERT INTO Patients (PatientID, Name, DateOfBirth) 
VALUES ('P001234', 'Fatima Ali', '1995-05-15');

INSERT INTO PatientPhones (PatientID, PhoneNumber) 
VALUES ('P001234', '0712345678');

-- إضافة زيارة تجريبية
INSERT INTO Visits (PatientID, DoctorID, ClinicCode, VisitDate) 
VALUES ('P001234', 'EMP01', 'CLN01', GETDATE());

-- إضافة تشخيص تجريبي
INSERT INTO Diagnoses (DiagnosisCode, Description) 
VALUES ('D101', 'Seasonal Influenza');