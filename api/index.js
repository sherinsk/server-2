const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');
const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Function to convert JSON to XLSX using ExcelJS
async function convertJsonToXlsx(students) {
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Students');

  // Define the columns for the worksheet
  worksheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Age', key: 'age', width: 10 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Phone Number', key: 'phoneNumber', width: 15 },
    { header: 'Address', key: 'address', width: 30 },
    { header: 'Date of Birth', key: 'dateOfBirth', width: 15 },
    { header: 'Enrollment Date', key: 'enrollmentDate', width: 15 },
    { header: 'Course', key: 'course', width: 20 },
    { header: 'Grade', key: 'grade', width: 10 },
    { header: 'Is Active', key: 'isActive', width: 10 },
    { header: 'Guardian Name', key: 'guardianName', width: 20 },
    { header: 'Guardian Phone', key: 'guardianPhone', width: 15 },
    { header: 'Gender', key: 'gender', width: 10 },
    { header: 'Nationality', key: 'nationality', width: 15 },
    { header: 'Profile Image URL', key: 'profileImageUrl', width: 30 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ];

  // Add the data into rows
  students.forEach(student => {
    worksheet.addRow({
      name: student.name,
      age: student.age,
      email: student.email,
      address: student.address,
      dateOfBirth: student.dateOfBirth,
      enrollmentDate: student.enrollmentDate,
      course: student.course,
      grade: student.grade,
      isActive: student.isActive,
      guardianName: student.guardianName,
      guardianPhone: student.guardianPhone,
      gender: student.gender,
      nationality: student.nationality,
      profileImageUrl: student.profileImageUrl,
      createdAt: student.createdAt
    });
  });

  // Write the workbook to a buffer and return it
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}

// API endpoint to download the XLSX file
app.post('/download-xlsx', async (req, res) => {
  try {
    // Fetch data from an external API
    const {allStudents}=req.body

    // Convert the fetched data to XLSX
    const xlsxData = await convertJsonToXlsx(allStudents);

    // Set the response headers for file download
    res.setHeader('Content-Disposition', 'attachment; filename=students.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Send the XLSX file as a response
    res.send(xlsxData);
  } catch (error) {
    console.error('Error fetching data from external API:', error);
    res.status(500).send('Failed to fetch data or create XLSX file.');
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
