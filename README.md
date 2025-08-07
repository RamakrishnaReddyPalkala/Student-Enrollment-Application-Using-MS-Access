# Student-Enrollment-Application-Using-MS-Access


# 🎓 Student Enrollment Management System (MS Access + VBA)

A lightweight **Student Enrollment Management System** built using **Microsoft Access** (2007–2016 format) with **VBA (Visual Basic for Applications)** to efficiently manage student data for academic institutions.

## 📌 Project Overview

This MS Access application provides a simple yet powerful interface for handling student records. It is designed to be used offline and offers automation features to streamline the data management process.

### ✅ Features

- 📁 Store and manage student details: Full Name, Gender, DOB, Contact, Course, and more
- 📝 Add and update student records via a user-friendly form
- 🔢 Auto-generate unique **Student IDs**
- 📄 Generate student reports with a single click
- ⚙️ Input validation powered by VBA to ensure data integrity
- 🖥️ Offline, standalone solution – no external server or internet required

---

## 🧱 Database Structure

### Table: `Students`

| Field Name     | Data Type   | Description                        |
|----------------|-------------|------------------------------------|
| `StudentID`    | AutoNumber  | Primary Key, Auto-generated        |
| `FullName`     | Short Text  | Student's full name                |
| `Gender`       | Short Text  | Gender (Male/Female)               |
| `DOB`          | Date/Time   | Date of Birth                      |
| `PhoneNumber`  | Short Text  | Contact Number                     |
| `Email`        | Short Text  | Email Address                      |
| `Course`       | Short Text  | Enrolled Course                    |
| `EnrolmentDate`| Date/Time   | Date of Enrollment                 |

---

## 🔧 Technology Stack

- **Microsoft Access (2007–2016 format)**
- **VBA (Visual Basic for Applications)** for logic and validations

---

## 🚀 Getting Started

1. Download or clone this repository.
2. Open the `.accdb` file in Microsoft Access (2007–2016).
3. Ensure macros are enabled to allow VBA scripts to run.
4. Start entering student records through the provided forms.


---

## 🔐 Data Validation

VBA scripts are used for:
- Validating required fields (e.g., Full Name, Phone Number, etc.)
- Ensuring correct data formats (e.g., valid email, numeric phone numbers)
- Preventing duplicate entries (optional feature)

---

## 📢 Note

- You can further enhance it by adding features like user login, course fee tracking, or export to Excel/PDF.

---

## 🧑‍💻 Author

**Ramakrishna Palkala**

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).
