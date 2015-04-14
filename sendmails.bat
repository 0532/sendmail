@echo off
set CLASSPATH=%CLASSPATH%;.;d:\sendmail\lib\spring.jar;d:\sendmail\lib\mail.jar;d:\sendmail\lib\servlet-api.jar;d:\sendmail\lib\dom4j-1.6.1.jar;d:\sendmail\lib\log4j-1.2.14.jar;d:\sendmail\lib\poi-3.7.jar;d:\sendmail\lib\poi-ooxml-3.7.jar;d:\sendmail\lib\poi-ooxml-schemas-3.7.jar;d:\sendmail\lib\xmlbeans-2.3.0.jar;
@echo Sending...
java SendMailGroup 
@echo ======================= off echo =======================
@echo. & pause