include VersionInfo

PKG_FILES=$(shell find R examples src man -not -path '*CVS*' -not -name '*~') DESCRIPTION

ifndef RHOME
RHOME=$(R_HOME)
endif

PACKAGE_DIR=/tmp
#$(RHOME)/src/library
INSTALL_DIR=${PACKAGE_DIR}/$(PKG_NAME)

include Install/GNUmakefile.admin

C_SOURCE=$(wildcard src/*.[ch] src/*.cpp) src/Makevars.win
R_SOURCE=$(wildcard R/*.[RS]) R/autoInterface.S
MAN_FILES=$(wildcard man/*.Rd)
INSTALL_DIRS=src man R
#RUN_TIME=inst/runtime/runtime.S

DOCS=Docs/overview.html Docs/overview.xml

R/autoInterface.S: R/autoInterface.xml
	$(MAKE) -C R $(@F)




package: DESCRIPTION  R/autoInterface.S $(DOCS)
	@if test -z "${RHOME}" ; then echo "You must specify RHOME" ; exit 1 ; fi
	if test -d $(INSTALL_DIR) ; then rm -fr $(INSTALL_DIR) ; fi
	mkdir $(INSTALL_DIR)
	cp NAMESPACE DESCRIPTION $(INSTALL_DIR)
	for i in $(INSTALL_DIRS) ; do \
	   mkdir $(INSTALL_DIR)/$$i ; \
	done
	cp -r $(C_SOURCE) $(INSTALL_DIR)/src
	cp -r $(MAN_FILES) $(INSTALL_DIR)/man
	cp -r $(R_SOURCE) $(INSTALL_DIR)/R
#	cp  install.R $(INSTALL_DIR)
	mkdir $(INSTALL_DIR)/inst
	mkdir $(INSTALL_DIR)/inst/Docs
	mkdir $(INSTALL_DIR)/inst/runtime
	if test -n "${DOCS}" ; then cp $(DOCS) $(INSTALL_DIR)/inst/Docs ; fi
	cp -r examples $(INSTALL_DIR)/inst
	cp  R/common.S $(INSTALL_DIR)/inst/runtime/
	find $(INSTALL_DIR) -name '*~' -exec rm {} \;
#	find $(INSTALL_DIR) -name 'CVS' -type d -exec rm -r {} \;


PWD=$(shell pwd)

release: source binary

binary:  package
	(cd $(RHOME)/src/library ; Rcmd build --binary $(PKG_NAME); mv $(ZIP_FILE) $(PWD))

zip:
	(cd $(RHOME)/library ; zip -r $(ZIP_FILE) $(PKG_NAME); mv $(ZIP_FILE) $(PWD))

tar source: package
	(cd $(INSTALL_DIR)/.. ; Rcmd build $(PKG_NAME); mv $(TAR_SRC_FILE) $(PWD))

install: package
	(cd $(RHOME)/src/gnuwin32 ; make pkg-$(PKG_NAME))

basicInstall: 
	(cd $(RHOME)/src/gnuwin32 ; make pkg-$(PKG_NAME))

check: package
	(cd $(RHOME)/src/library ; Rcmd check $(PKG_NAME))

file:
	@echo "${PKG_FILES}"


Docs/%.html: Docs/%.xml
	$(MAKE) -C Docs $(@F)
