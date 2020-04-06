#!/bin/bash

EXEL_FILE='/c/TSDE_Workarea/NHI5HC/hieunguyen/0000_Project/Management/Jira_Task_List.xlsm'
XLSX2CSV='/c/Python27/python.exe /c/Python27/Lib/site-packages/xlsx2csv.py'

#### Create List User #######
${XLSX2CSV} -s 2 ${EXEL_FILE} | grep 'banvien' | awk -F, '{print $2}' | sed 's/@.*$//g' > ./LIST

array="ASW ASW/ESDL ASW/Param ASW/Param/Param1D ASW/Param/Param2D"

for folder in `echo $array`
do
  echo "Create ${folder}"
  mkdir -p ./${folder}
  if [ `echo ${folder} | egrep -c "ESDL|Param1D|Param2D"` -eq 1 ]
  then
    for user in `cat ./LIST`
    do
      echo "Create folder ${user} at folder ${folder}"
      mkdir -p ./${folder}/${user}
    done
  fi
done

