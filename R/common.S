setMethod("createTypeVarName",
            "ClassDefinition",
# Map the given names in var to a unique and legitimate
# R variable name for the given class.
# 
function(className, var, quote = TRUE)
{
  createTypeVarName(className@className, var, quote)
})
	