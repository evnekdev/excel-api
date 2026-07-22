# Excel errors

The semantic error value stores the exact signed physical `VT_ERROR.scode`. The conventional mappings are `#NULL! 0x800A07D0`, `#DIV/0! 0x800A07D7`, `#VALUE! 0x800A07DF`, `#REF! 0x800A07E7`, `#NAME? 0x800A07ED`, `#NUM! 0x800A07F4`, and `#N/A 0x800A07FA`. Unknown negative SCODEs round-trip; short worksheet numbers such as 2042 are rejected for direct raw writes.
