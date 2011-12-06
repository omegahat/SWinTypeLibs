LOCAL_EXAMPLES=1
OMEGA_HOME=../../..

VERSION_INFO=../VersionInfo

include $(OMEGA_HOME)/Config/GNUmakefile.SWebPackageConfig
include $(OMEGA_HOME)/R/Config/RPackageWebConfig


CSS_FILES=$(OMEGA_HOME)/Docs/OmegaTech.css

TODO=../Todo.html
CHANGES=../Changes.html
AUX_DOCS=$(BUGS) $(TODO) $(FAQ) $(CHANGES)

OTHER_DOCS=../INSTALL.html ../README.html

include $(OMEGA_HOME)/Config/GNUmakefile.SWebPackageRules

basicInstall: index.html 
	scp -r index.html $(OTHER_DOCS) $(CSS_FILES) $(AUX_DOCS) $(TARGET_DIR)


ifdef LOCAL_EXAMPLES

EXAMPLES=$(shell find ../examples -not -name '*~' -not -path  '*/CVS/*')

examples: 
	scp $(EXAMPLES) $(TARGET_DIR)/examples

endif

install: pkg/$(PKG_NAME)/inst/Docs index.html examples ../Todo.html
	scp -r index.html examples  $(OTHER_DOCS) pkg/$(PKG_NAME)/inst/Docs $(CSS_FILES) $(AUX_DOCS) $(PKG_TAR_FILE) $(TARGET_DIR)
	$(update-r-mirror)

