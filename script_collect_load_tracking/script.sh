#!/bin/bash -x

#set +x
you_want_create_list_database_again="YES"
you_want_update_list_check_again="YES"

#Create your list finding from database
list_group_package=",20200511,"
your_input_path="$1"

XLSX2CSV="/c/Python27/python.exe /c/Python27/Lib/site-packages/xlsx2csv.py "
SUMMARY="/c/Users/nhi5hc/Desktop/Book1.xlsx"

SHEET_NAME="Merged_COEM"

COL_No=""
COL_Project=
#COL_ComponentName=""
#COL_ItemName=""
#COL_Database=""
COL_Tester=""
COL_ELOC_Recheck_by_Tool=""
COL_Estimation_base_TOOL=""
COL_Actual_hour=""
COL_Status=""
COL_Planned_End=""

if [ "${you_want_update_list_check_again}" == "YES" ]
then
  echo "CREATE LIST CHECK AGAIN"
  ${XLSX2CSV} -n ${SHEET_NAME} ${SUMMARY} | sed -n '/^,No,Package/,/^,Table KPI ASW/p' | egrep -v '^,\+$' | sed 's/^,\+//g' > .TEMP_SUMMARY

  COL_No=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "No") {print i; break}; i++}}'`
  COL_Project=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Project") {print i; break}; i++}}'`
  #COL_ComponentName=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "ComponentName") {print i; break}; i++}}'`
  #COL_ItemName=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "ItemName") {print i; break}; i++}}'`
  #COL_Database=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Database") {print i; break}; i++}}'`
  COL_Tester=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Tester") {print i; break}; i++}}'`
  COL_ELOC_Recheck_by_Tool=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "ELOC Recheck by Tool") {print i; break}; i++}}'`
  COL_Estimation_base_TOOL=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Estimation base TOOL") {print i; break}; i++}}'`
  COL_Actual_hour=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Actual hour") {print i; break}; i++}}'`
  COL_Status=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Status") {print i; break}; i++}}'`
  COL_Planned_End=`grep "^No," .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Planned End") {print i; break}; i++}}'`

  grep '^[0-9]\+,' .TEMP_SUMMARY \
    | grep -v '^,No,' \
    | awk -v col_no=$COL_No -v col_project=$COL_Project -v col_tester=$COL_Tester -v col_eloc_recheck_by_tool=$COL_ELOC_Recheck_by_Tool -v col_estimation_base_tool=$COL_Estimation_base_TOOL \
          -v col_actual_hour=$COL_Actual_hour -v col_status=$COL_Status -v col_planned_end=$COL_Planned_End \
          -F, '{printf "%s,%s,%s,%s,%s,%s,%s,%s\n", \
               $col_no, $col_project, $col_tester, $col_eloc_recheck_by_tool, $col_estimation_base_tool, $col_actual_hour, $col_status, $col_planned_end}' \
    | sed 's|\\|/|g' | sed 's/\s*\s/ /g' > ./LIST
fi

COL_No=1
COL_Project=2
#COL_ComponentName=""
#COL_ItemName=""
#COL_Database=""
COL_Tester=3
COL_ELOC_Recheck_by_Tool=4
COL_Estimation_base_TOOL=5
COL_Actual_hour=6
COL_Status=7
COL_Planned_End=8


rm -rf LOG.txt
touch LOG.txt

for user in `cat ./LIST | grep "$prj" | awk -v col_tester=$COL_Tester -F, '{printf "%s\n", $col_tester}' | sort -u`
do
    for prj in "`cat ./LIST | awk -v col_project=$COL_Project -F, '{printf "%s\n", $col_project}' | sort -u`"
    do
        grep ",$prj," ./LIST | grep ",$user," >> LOG.txt
    done

done

