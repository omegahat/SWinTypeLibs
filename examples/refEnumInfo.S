library(SWinTypeLibs)

# Load the Excel type library.
lib = LoadTypeLib("C:/Microsoft office/OFFICE11/Excel.exe")

# Deal with the ShapeNodes element in the type library 
# as it has a SetSegmentType() method that has a parameter
# of a type that is not in the Excel type library.
#
info = lib[["ShapeNodes"]]

 # Get the descriptions of the functions and the SetSegmentType() specifically.
funcs = getFuncs(info, lib = lib)
func = funcs$SetSegmentType

# Now get its SegmentType method. 
parm = func@parameters$SegmentType

# Now, we can get its 
tinfo = getRefTypeInfo(info, parm@type)
name = getRefTypeName(info, parm@type)

olib = getTypeLib(tinfo)

# office = LoadTypeLib("C:/Program Files/Common Files/Microsoft Shared/OFFICE11/MSO.dll")

if(!SWinTypeLibs:::isSame(lib, olib)) {
    
}

