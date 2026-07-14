/* Prepare the [class] sheet for import: wrap names in %...%, then strip them.

   This is the portable-SAS core of the utl-altair-slc-excel recipe: the input
   DATA step that builds the sheet (name=cats('%',name,'%')) and the "remove the
   % characters" transformation the whole repo is about. The upstream recipe does
   the strip through an Excel/ODBC pass-through (replace(name,'%','')); here the
   same result is produced with base SAS so it runs anywhere, and the $hex.
   verification columns from the recipe are kept so you can see the byte layout.
*/

/* --- input: build the sheet exactly as the recipe does ------------------- */
data class;
  length name $8;
  input
    name$
    sex$ age;
    name=cats('%',name,'%');
cards4;
Alfred  M 14
Alice   F 13
Barbara F 13
Carol   F 14
Henry   M 14
James   M 12
;;;;
run;

/* --- prepare for import: strip the % wrappers, keep a hex view ----------- */
data tab;
  set class;
  length clean_name $8;
  clean_name    = compress(name, '%');   /* recipe's replace(name,'%','')   */
  name_hex      = put(name,       $hex18.);
  clean_name_hex= put(clean_name, $hex14.);
run;

proc print data=tab noobs;
  var name name_hex sex clean_name_hex age clean_name;
run;
