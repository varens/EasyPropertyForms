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
  [assigned]      INTEGER CONSTRAINT fk_personnel_id REFERENCES tbl_personnel ([id])
);
