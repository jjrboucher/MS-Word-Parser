/*
List of files where a specific user is the author
In this example, the user's name is "user"
*/

SELECT file_name
FROM people
WHERE author == "user"
/* if you want to search for more than one author, replace the WHERE statement above with the following
WHERE author IN ("user", "user1", "user2", etc.)
*/