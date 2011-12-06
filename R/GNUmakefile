XSL_DIR=$(OMEGA_HOME)/XML/Literate

include $(OMEGA_HOME)/Config/GNUmakefile.xml

autoInterface.S: autoInterface.xml
	xsltproc -o $@ -param language "'S'" -param group "'def'" $(XSL_DIR)/croutine.xsl autoInterface.xml