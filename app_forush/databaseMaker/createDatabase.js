const ADODB = require("node-adodb");
const { Sequelize, DataTypes } = require("sequelize");

async function run() {
  // Connect to the Access database
  const path = process.cwd() + "/database.mdb";
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

  // Retrieve the rows from the table
  const rows = await connection.query("SELECT * FROM Foroush_Detail");

  // Reverse the rows
  const all = rows.reverse();

  // Insert each row into the ORM
  for (const row of all) {
    console.log(row.Radif);
    try {
      await MyTable.create({
        KharidarName: row.KharidarName,
        KharidarLastNameSherkatName: row.KharidarLastNameSherkatName,
        KharidarNationalCode: row.KharidarNationalCode,
        HCKharidarTypeCode: row.HCKharidarTypeCode,
      });
    } catch (err) {}
  }

  // Close the database connection
  connection.close();
}

run();
