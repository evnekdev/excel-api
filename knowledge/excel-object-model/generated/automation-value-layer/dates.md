# OA dates

`OaDate` stores a checked finite serial. `DateVariant` encodes a non-negative serial as `VT_DATE`; a negative serial is rejected. `Value2Serial` explicitly encodes the serial as `VT_R8` and is the supported policy for negative OA numeric storage.
