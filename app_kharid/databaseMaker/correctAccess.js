const ADODB = require("node-adodb");
const { Sequelize, DataTypes } = require("sequelize");

async function run() {
  // Connect to the Access database
  const path = "C:\\Users\\a\\Desktop\\safar\\TTMS.mdb";
  const connStr = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${path};`;

  const connection = ADODB.open(connStr);

  // Define the model for the table
  const sequelize = new Sequelize({
    dialect: "sqlite",
    storage: "mydatabase.db",
  });

  const MyTable = sequelize.define(
    "companies",
    {
      KharidarName: DataTypes.STRING,
      KharidarLastNameSherkatName: DataTypes.STRING,
      KharidarNationalCode: DataTypes.STRING,
      HCKharidarTypeCode: DataTypes.STRING,
    },
    { timestamps: false }
  );

  // Retrieve all the records from the ORM
  const companies = await MyTable.findAll();

  let i = 0;
  for (const company of companies) {
    console.log(i);

    if (company.KharidarName) {
      const sql_update = `UPDATE Foroush_Detail SET HCKharidarTypeCode = '${company.HCKharidarTypeCode}', KharidarNationalCode = '${company.KharidarNationalCode}', KharidarLastNameSherkatName = '${company.KharidarLastNameSherkatName}' WHERE KharidarName = '${company.KharidarName}'`;
      connection.execute(sql_update);
    }

    if (company.KharidarLastNameSherkatName) {
      const sql_update = `UPDATE Foroush_Detail SET HCKharidarTypeCode = '${company.HCKharidarTypeCode}', KharidarNationalCode = '${company.KharidarNationalCode}', KharidarName = '${company.KharidarName}' WHERE KharidarLastNameSherkatName = '${company.KharidarLastNameSherkatName}'`;
      connection.execute(sql_update);
    }

    if (company.KharidarNationalCode) {
      const sql_update = `UPDATE Foroush_Detail SET HCKharidarTypeCode = '${company.HCKharidarTypeCode}', KharidarName = '${company.KharidarName}', KharidarLastNameSherkatName = '${company.KharidarName}' WHERE KharidarNationalCode = '${company.KharidarNationalCode}'`;
      connection.execute(sql_update);
    }

    i++;
  }

  // Close the database connection
  connection.close();
}

run();
