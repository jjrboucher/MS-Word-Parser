/*
Example of a query to obtain a list of files with a given rsidR value 
as identifed in another query. If you are seeing the same rsidR value across different files
with different rsidRoot values, that could be a collision (improbable), or it could mean those files
were created from a common template (probable). On a Windows system, the template is specific to each 
Windows user account.
*/

SELECT r.file_name, count(r.rsid_value), m.rsid_root, m.author, m.template
FROM rsids AS r
LEFT JOIN metadata AS m ON m.file_name == r.file_name
WHERE r.rsid_value == "008C0F1F"
/* to query multiple values, you could use: WHERE r.rsid_value IN ("value1", "value2", etc.) */
GROUP BY r.file_name
ORDER BY m.rsid_root