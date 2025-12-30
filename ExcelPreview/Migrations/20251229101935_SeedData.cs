using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional

namespace Backend.Migrations
{
    /// <inheritdoc />
    public partial class SeedData : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ExcelDatas",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Name = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    Value = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelDatas", x => x.Id);
                });

            migrationBuilder.InsertData(
                table: "ExcelDatas",
                columns: new[] { "Id", "Name", "Value" },
                values: new object[,]
                {
                    { 1, "01KDMSXN7FV3S9FK4TRNXZHP8Z", "67" },
                    { 2, "01KDMSXN7FAKWQTMFRPHP52PAX", "15" },
                    { 3, "01KDMSXN7FHA3JQEDK7W97W71N", "13" },
                    { 4, "01KDMSXN7F7MDQZKC6XNCBP3R4", "53" },
                    { 5, "01KDMSXN7FRR1FCPPQ7T7ZMPFA", "17" },
                    { 6, "01KDMSXN7FWTE92R14CM2D0RN5", "27" },
                    { 7, "01KDMSXN7F0PHDAZTKHWAMTR8M", "73" },
                    { 8, "01KDMSXN7F3ZHTJTRPBC4BY82E", "52" },
                    { 9, "01KDMSXN7FYTQ9RPWFDETQ2TBP", "18" },
                    { 10, "01KDMSXN7FTT2A3X5XPZBP331T", "77" }
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ExcelDatas");
        }
    }
}
