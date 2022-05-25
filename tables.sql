CREATE TABLE tbl_personnel
(
  [id]            COUNTER CONSTRAINT pk_personnel_id PRIMARY KEY,
  [name]          CHAR,
  [rank]          CHAR,
  [unit]          CHAR,
  [designation]   CHAR
);
=sep=
CREATE TABLE tbl_property
(
  [id]            COUNTER CONSTRAINT pk_property_id PRIMARY KEY,
  [type]          CHAR,
  [nomenclature]  CHAR,
  [nsn]           CHAR,
  [lin]           CHAR,
  [serial]        CHAR,
  [admin]         CHAR,
  [assigned]      INTEGER CONSTRAINT fk_personnel_id REFERENCES tbl_personnel ([id]),
  [toggle]        BIT
);
=sep=
SELECT tbl_personnel.[id] AS personnel_id,
  tbl_property.[id] AS property_id,
  tbl_property.[toggle],
  tbl_property.[type],
  tbl_property.[nomenclature],
  tbl_property.[nsn],
  tbl_property.[lin],
  tbl_property.[serial],
  tbl_property.[admin],
  tbl_property.[assigned]
FROM tbl_property
  LEFT JOIN tbl_personnel ON tbl_property.assigned=tbl_personnel.id;
=sep=
SELECT [id], [rank] & ' ' & [name] AS person
FROM tbl_personnel
ORDER BY [name]
=sep=
tbl_personnel.[rank] & ' ' & tbl_personnel.[name] AS person,