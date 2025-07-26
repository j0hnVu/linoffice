#!/bin/bash

# this script still needs to be written. it's supposed to take the values from registry_override.conf and apply them to the Windows registry

# could do a variant of the locale_reg.sh that then creates a .reg file with the appropriate values and then calls freerdp to add this to the Windows registry 

# DATE_FORMAT="" ---> the separator (. / -) goes into sDate and the date/month/year order goes in sShortDate e.g. dd/MM/yyyy
# DECIMAL_SEPARATOR="" --> sDecimal & sMonDecimalSep; then the opposite (, or .) for sThousand and sMonThousandSep
# CURRENCY_SYMBOL="" ---> apply in sCurrency (need to make sure encoding is correct)

# Should also add something to locale_reg.sh to prefill the 3 values in the registry_override.sh and then change the mainwindow.py to read these values

