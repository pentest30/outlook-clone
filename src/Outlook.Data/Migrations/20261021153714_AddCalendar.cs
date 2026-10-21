using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Outlook.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddCalendar : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "CalendarIcs",
                table: "Messages",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "HasCalendar",
                table: "Messages",
                type: "INTEGER",
                nullable: false,
                defaultValue: false);

            migrationBuilder.CreateTable(
                name: "Appointments",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    AccountId = table.Column<int>(type: "INTEGER", nullable: false),
                    Uid = table.Column<string>(type: "TEXT", nullable: false),
                    Sequence = table.Column<int>(type: "INTEGER", nullable: false),
                    Title = table.Column<string>(type: "TEXT", nullable: false),
                    Location = table.Column<string>(type: "TEXT", nullable: false),
                    StartUtc = table.Column<long>(type: "INTEGER", nullable: false),
                    EndUtc = table.Column<long>(type: "INTEGER", nullable: false),
                    IsAllDay = table.Column<bool>(type: "INTEGER", nullable: false),
                    IsRecurring = table.Column<bool>(type: "INTEGER", nullable: false),
                    SeriesEndUtc = table.Column<long>(type: "INTEGER", nullable: true),
                    OrganizerEmail = table.Column<string>(type: "TEXT", nullable: false),
                    OrganizerName = table.Column<string>(type: "TEXT", nullable: false),
                    IsOrganizer = table.Column<bool>(type: "INTEGER", nullable: false),
                    Attendees = table.Column<string>(type: "TEXT", nullable: false),
                    MyResponse = table.Column<int>(type: "INTEGER", nullable: false),
                    ShowAs = table.Column<int>(type: "INTEGER", nullable: false),
                    ReminderMinutes = table.Column<int>(type: "INTEGER", nullable: true),
                    RemindersDismissedThroughUtc = table.Column<long>(type: "INTEGER", nullable: true),
                    Ics = table.Column<string>(type: "TEXT", nullable: false),
                    LastModifiedUtc = table.Column<long>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Appointments", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Appointments_Accounts_AccountId",
                        column: x => x.AccountId,
                        principalTable: "Accounts",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Appointments_AccountId_StartUtc",
                table: "Appointments",
                columns: new[] { "AccountId", "StartUtc" });

            migrationBuilder.CreateIndex(
                name: "IX_Appointments_AccountId_Uid",
                table: "Appointments",
                columns: new[] { "AccountId", "Uid" },
                unique: true);

            // Bodies cached before this version were read without looking for calendar parts, so
            // fetch them again on next open (it is a cache; the server still has everything).
            migrationBuilder.Sql("UPDATE Messages SET IsBodyDownloaded = 0, BodyText = NULL, BodyHtml = NULL;");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Appointments");

            migrationBuilder.DropColumn(
                name: "CalendarIcs",
                table: "Messages");

            migrationBuilder.DropColumn(
                name: "HasCalendar",
                table: "Messages");
        }
    }
}
