using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Outlook.Data.Migrations
{
    /// <inheritdoc />
    public partial class AddSearchAndCategories : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "CategoryKeywords",
                table: "Messages",
                type: "TEXT",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "SearchText",
                table: "Messages",
                type: "TEXT",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "Categories",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    AccountId = table.Column<int>(type: "INTEGER", nullable: false),
                    Name = table.Column<string>(type: "TEXT", nullable: false),
                    Color = table.Column<string>(type: "TEXT", nullable: false),
                    Keyword = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Categories", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Categories_Accounts_AccountId",
                        column: x => x.AccountId,
                        principalTable: "Accounts",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Categories_AccountId_Keyword",
                table: "Categories",
                columns: new[] { "AccountId", "Keyword" },
                unique: true);

            // Full-text index over the searchable columns. "External content" mode: the text lives
            // only in Messages, and these triggers keep the index in step with it. Flag changes
            // (read, flagged, moved) do not touch indexed columns, so they cost nothing here.
            migrationBuilder.Sql("""
                UPDATE Messages SET SearchText = BodyText WHERE BodyText IS NOT NULL;

                CREATE VIRTUAL TABLE MessageSearch USING fts5(
                    Subject, FromName, FromAddress, "To", Cc, Preview, SearchText,
                    content='Messages', content_rowid='Id', tokenize='unicode61 remove_diacritics 2');

                CREATE TRIGGER Messages_SearchInsert AFTER INSERT ON Messages BEGIN
                    INSERT INTO MessageSearch(rowid, Subject, FromName, FromAddress, "To", Cc, Preview, SearchText)
                    VALUES (new.Id, new.Subject, new.FromName, new.FromAddress, new."To", new.Cc, new.Preview, new.SearchText);
                END;

                CREATE TRIGGER Messages_SearchDelete AFTER DELETE ON Messages BEGIN
                    INSERT INTO MessageSearch(MessageSearch, rowid, Subject, FromName, FromAddress, "To", Cc, Preview, SearchText)
                    VALUES ('delete', old.Id, old.Subject, old.FromName, old.FromAddress, old."To", old.Cc, old.Preview, old.SearchText);
                END;

                CREATE TRIGGER Messages_SearchUpdate AFTER UPDATE OF Subject, FromName, FromAddress, "To", Cc, Preview, SearchText ON Messages BEGIN
                    INSERT INTO MessageSearch(MessageSearch, rowid, Subject, FromName, FromAddress, "To", Cc, Preview, SearchText)
                    VALUES ('delete', old.Id, old.Subject, old.FromName, old.FromAddress, old."To", old.Cc, old.Preview, old.SearchText);
                    INSERT INTO MessageSearch(rowid, Subject, FromName, FromAddress, "To", Cc, Preview, SearchText)
                    VALUES (new.Id, new.Subject, new.FromName, new.FromAddress, new."To", new.Cc, new.Preview, new.SearchText);
                END;

                INSERT INTO MessageSearch(MessageSearch) VALUES ('rebuild');
                """);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql("""
                DROP TRIGGER IF EXISTS Messages_SearchUpdate;
                DROP TRIGGER IF EXISTS Messages_SearchDelete;
                DROP TRIGGER IF EXISTS Messages_SearchInsert;
                DROP TABLE IF EXISTS MessageSearch;
                """);

            migrationBuilder.DropTable(
                name: "Categories");

            migrationBuilder.DropColumn(
                name: "CategoryKeywords",
                table: "Messages");

            migrationBuilder.DropColumn(
                name: "SearchText",
                table: "Messages");
        }
    }
}
