using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Outlook.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddAddressBook : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "KnownAddresses",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    AccountId = table.Column<int>(type: "INTEGER", nullable: false),
                    Address = table.Column<string>(type: "TEXT", nullable: false),
                    Name = table.Column<string>(type: "TEXT", nullable: false),
                    SentCount = table.Column<int>(type: "INTEGER", nullable: false),
                    ReceivedCount = table.Column<int>(type: "INTEGER", nullable: false),
                    LastUsedUtc = table.Column<long>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_KnownAddresses", x => x.Id);
                    table.ForeignKey(
                        name: "FK_KnownAddresses_Accounts_AccountId",
                        column: x => x.AccountId,
                        principalTable: "Accounts",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_KnownAddresses_AccountId_Address",
                table: "KnownAddresses",
                columns: new[] { "AccountId", "Address" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "KnownAddresses");
        }
    }
}
