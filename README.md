#COUNTU_UDF Excel Module

This COUNTU_UDF module adds a set of user defined functions for Excel which add the ability to count the number of items in a given range against as set of criteria whilst ignoring duplicates.

* COUNTU
* COUNTUIF
* COUNTUIFS

**The module requires a reference to the Microsoft Scripting Runtime library.**

##Functions

###COUNTU

Counts the number of cells in a range excluding any duplicates
found.

**Usage:**
```
=COUNTU(range)
```

###COUNTUIF

Counts the number of cells in a range that meet the given
condition excluding any duplicates found in the
duplicate_range.

**Usage:**
```
=COUNTUIF(duplicate_range, criteria, criteria_range)
```

###COUNTUIFS

Counts the number of cells specified by a given set of
conditions or criteria excluding any duplicates found in the
duplicate_range.

**Usage:**
```
=COUNTUIFS(duplicate_range, criteria_range1, criteria1, …)
```
