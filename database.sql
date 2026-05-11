-- SkillTrack System — Database Schema
-- Run this file in your MySQL client to set up the demo database

CREATE DATABASE IF NOT EXISTS skilltrack;
USE skilltrack;

-- 1. Learners
CREATE TABLE IF NOT EXISTS learners (
    learner_id   INT AUTO_INCREMENT PRIMARY KEY,
    full_name    VARCHAR(100) NOT NULL,
    email        VARCHAR(150) UNIQUE NOT NULL,
    phone        VARCHAR(20),
    joined_date  DATE DEFAULT (CURRENT_DATE)
);

-- 2. Instructors
CREATE TABLE IF NOT EXISTS instructors (
    instructor_id    INT AUTO_INCREMENT PRIMARY KEY,
    full_name        VARCHAR(100) NOT NULL,
    email            VARCHAR(150) UNIQUE NOT NULL,
    specialization   VARCHAR(100)
);

-- 3. Courses
CREATE TABLE IF NOT EXISTS courses (
    course_id      INT AUTO_INCREMENT PRIMARY KEY,
    title          VARCHAR(200) NOT NULL,
    description    TEXT,
    start_date     DATE,
    end_date       DATE,
    instructor_id  INT,
    FOREIGN KEY (instructor_id) REFERENCES instructors(instructor_id) ON DELETE SET NULL
);

-- 4. Enrollments
CREATE TABLE IF NOT EXISTS enrollments (
    enrollment_id  INT AUTO_INCREMENT PRIMARY KEY,
    learner_id     INT NOT NULL,
    course_id      INT NOT NULL,
    enrolled_date  DATE DEFAULT (CURRENT_DATE),
    FOREIGN KEY (learner_id) REFERENCES learners(learner_id) ON DELETE CASCADE,
    FOREIGN KEY (course_id)  REFERENCES courses(course_id)  ON DELETE CASCADE
);

-- 5. Progress
CREATE TABLE IF NOT EXISTS progress (
    progress_id    INT AUTO_INCREMENT PRIMARY KEY,
    enrollment_id  INT NOT NULL,
    completion_pct DECIMAL(5,2) DEFAULT 0.00,
    last_updated   DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (enrollment_id) REFERENCES enrollments(enrollment_id) ON DELETE CASCADE
);

-- 6. Certificates
CREATE TABLE IF NOT EXISTS certificates (
    certificate_id  INT AUTO_INCREMENT PRIMARY KEY,
    enrollment_id   INT NOT NULL,
    issued_date     DATE DEFAULT (CURRENT_DATE),
    grade           VARCHAR(10),
    FOREIGN KEY (enrollment_id) REFERENCES enrollments(enrollment_id) ON DELETE CASCADE
);

-- 7. Attendance
CREATE TABLE IF NOT EXISTS attendance (
    attendance_id  INT AUTO_INCREMENT PRIMARY KEY,
    enrollment_id  INT NOT NULL,
    attend_date    DATE NOT NULL,
    status         ENUM('present','absent','late') DEFAULT 'present',
    FOREIGN KEY (enrollment_id) REFERENCES enrollments(enrollment_id) ON DELETE CASCADE
);

-- ─── Demo Data ───────────────────────────────────────────────────────────────

INSERT INTO instructors (full_name, email, specialization) VALUES
('Ahmed Hassan',  'ahmed@skilltrack.io',  'Web Development'),
('Sara Mohamed',  'sara@skilltrack.io',   'Data Science'),
('Omar Khalil',   'omar@skilltrack.io',   'Cybersecurity');

INSERT INTO learners (full_name, email, phone, joined_date) VALUES
('Ali Fathy',      'ali@mail.com',    '01012345678', '2024-01-10'),
('Nour Adel',      'nour@mail.com',   '01098765432', '2024-02-15'),
('Khaled Samir',   'khaled@mail.com', '01123456789', '2024-03-01'),
('Hana Ibrahim',   'hana@mail.com',   '01234567890', '2024-03-20');

INSERT INTO courses (title, description, start_date, end_date, instructor_id) VALUES
('Python Basics',       'Intro to Python programming',        '2024-04-01', '2024-06-30', 1),
('Machine Learning 101','Practical ML with scikit-learn',     '2024-04-15', '2024-07-15', 2),
('Network Security',    'Fundamentals of ethical hacking',    '2024-05-01', '2024-08-01', 3);

INSERT INTO enrollments (learner_id, course_id, enrolled_date) VALUES
(1, 1, '2024-04-01'),
(2, 1, '2024-04-02'),
(3, 2, '2024-04-16'),
(4, 3, '2024-05-02'),
(1, 2, '2024-04-20');

INSERT INTO progress (enrollment_id, completion_pct) VALUES
(1, 80.00),
(2, 55.50),
(3, 30.00),
(4, 95.00),
(5, 10.00);

INSERT INTO certificates (enrollment_id, issued_date, grade) VALUES
(4, '2024-07-30', 'A');

INSERT INTO attendance (enrollment_id, attend_date, status) VALUES
(1, '2024-04-01', 'present'),
(1, '2024-04-08', 'late'),
(2, '2024-04-01', 'present'),
(3, '2024-04-16', 'absent'),
(4, '2024-05-02', 'present'),
(4, '2024-05-09', 'present');
