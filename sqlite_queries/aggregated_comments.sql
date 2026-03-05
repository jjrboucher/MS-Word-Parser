SELECT  DISTINCT C.file_name,
        C.author,
        C.initials,
        C.timestamp_utc,
        C.comment_id,
        C.comment_paraid,
        C.paraid_text,
        EC.paraidparent,
CASE EC.done
WHEN 0 THEN "FALSE"
WHEN 1 THEN "TRUE"
END AS "Done?",
CID.durableId,
EC2.dateUtc,
EC2.reactionType,
EC2.reactionDateUtc,
EC2.uri,
EC2.userId,
EC2.userProvider,
EC2.userName
FROM comments AS C
    LEFT JOIN (SELECT file_name, paraid, paraidparent, done FROM extended_comments) AS EC ON C.paraid_text == EC.paraId AND C.file_name == EC.file_name
LEFT JOIN (SELECT file_name, paraid, durableid FROM Comments_ids) AS CID ON C.comment_paraid == CID.paraId AND C.file_name == CID.file_name
LEFT JOIN (SELECT file_name, durableId, dateUtc, reactionType, reactionDateUtc, uri, userId, userProvider, userName FROM "Extensible Comments") AS EC2 ON CID.durableId == EC2.durableId AND CID."File Name" == EC2."File Name"