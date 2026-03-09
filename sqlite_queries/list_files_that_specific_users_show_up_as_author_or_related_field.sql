/*
Search for files that a select users show up as author or other related field.

The SQLite statement will need to be modified if some of the tables do not exist. For example, if nobody
reacted to a comment (e.g., Liked a comment), there will be no extensible_comments table.

Simply delete the UNION statement for tables that do not exist.

From Claude.ai
Potential improvements:
If you query this frequently, indexes on the searched columns will make a significant difference:
CREATE INDEX IF NOT EXISTS idx_metadata_author ON metadata(author);
CREATE INDEX IF NOT EXISTS idx_metadata_lmb ON metadata(last_modified_by);
CREATE INDEX IF NOT EXISTS idx_comments_author ON comments(author);
CREATE INDEX IF NOT EXISTS idx_ext_comments_user ON extensible_comments(username);
CREATE INDEX IF NOT EXISTS idx_people_author ON people(author);

Claude.ai provided the following option to use a variable (kind of) to store the authors rather than repeating them each time.
The statement is not as legible. But the advantage being that you only define the authors at one location, then reference them
in each statement. Making modifications to the statement to query for other/more authors only requires us to change that in the WITH statement.

WITH target_users AS (
    SELECT '<name_1>' AS username
    UNION ALL SELECT '<name_2>'
)
SELECT file_name FROM metadata
WHERE author IN (SELECT username FROM target_users)
   OR last_modified_by IN (SELECT username FROM target_users)
UNION
SELECT file_name FROM comments
WHERE author IN (SELECT username FROM target_users)
UNION
SELECT file_name FROM extensible_comments
WHERE username IN (SELECT username FROM target_users)
UNION
SELECT file_name FROM people
WHERE author IN (SELECT username FROM target_users)
*/
SELECT file_name
FROM metadata
WHERE author IN ("<name_1>", "<name_2>") OR last_modified_by IN ("<name_1>", "<name_2>")

UNION

SELECT file_name
FROM comments
WHERE author IN ("<name_1>", "<name_2>")

UNION

SELECT file_name
FROM extensible_comments
WHERE username IN ("<name_1>", "<name_2>")

UNION

SELECT file_name
FROM people
WHERE author IN ("<name_1>", "<name_2>")