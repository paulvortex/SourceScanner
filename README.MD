SourceScanner
----

This little utility written in VB builds a .DEF/.ENT files used 
by GtkRadiant directly from mod's source (finding /*QUAKED entries)

Features
------

- Extracting of /*QUAKED entity definitions from source files as .def file (Quake entity definitions).
- Extracting of /*GTKRAD15 entity definitions from source files as .ent file (GtkRadiant 1.5-introduced XML entity definitions).
- Extracting of /*SS definitions into either .def or .ent
- Entity groups that holds common keys/flags can be set in external ini file which should be located in: [destination folder]\ssconfig.ini or [destination folder]\..\[destination folder].ini

--------------------------------------------------------------------------------
 Version History + Changelog
--------------------------------------------------------------------------------

1.0 (16-07-2013)
------
- initial commit (borrowed from archives)