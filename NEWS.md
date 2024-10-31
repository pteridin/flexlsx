# flexlsx 0.3.0

* Numerics will be written as numerics unless 
`options("openxlsx2.string_nums" = TRUE)` is set - thanks to @JanMarvin
* BUGFIX #28: flextables without headers will be correctly displayed

# flexlsx 0.2.2

* CRAN release :)
* Added Bugfix: Can't export flextable with no header (#28)

# flexlsx 0.2.1

* Fixes for CRAN release

# flexlsx 0.1.3

* Release candidate for CRAN

# flexlsx 0.1.2

* Column width will now be set (#3)
* Row height will now be set (#11)
* Bugfixes:
  * Number stored as text warning removed (#5)
  * Text-colors will now be handled correct (maybe?, #4)
  * Caption is longer than table (#3)
  * Caption `<br>` should be replaced by `\n` and the cell be wrapped? (#3)
  * Merged cells border issues fixed (#3)

# flexlsx 0.1.1

* Several Bugfixes regarding caption generation
* Cleaned up `devtools::check()` warnings & documentation errors

# flexlsx 0.1.0

* First implementation
