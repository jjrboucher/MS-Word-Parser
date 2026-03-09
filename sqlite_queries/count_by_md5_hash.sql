/*
Simple query to get a count of how many files there are for a given md5_hash value
And only return hash values where the hash matches more than one file.
*/
SELECT md5_hash, COUNT(*) AS "total count"
FROM document_summary
GROUP BY md5_hash
HAVING COUNT(*) > 1
ORDER BY "total count" DESC