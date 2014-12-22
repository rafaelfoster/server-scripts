#!/bin/bash
#
# This script get all Mysql Backup size (except the internal one) and display it as a Json
# that can be used to Zabbix Discovery Rules.
#
# This simple script was created by RafaelFoster: rafaelgfoster (at) gmail (dot) com
#
# Json Output: {"data": [ { "{#DBNAME}" : "database_name","{#DBSIZE}":"size_in_bytes"},{"{#DBNAME}":"database_name2","{#DBSIZE}":"size_in_bytes"} ] }

bkp_folder="path_to_backups"
cd "$bkp_folder"

db_name=$1
if [ ! $# -gt 0 ]; then
   db_name="BKP.DISCOVERY"
fi

IFSBKP=$IFS
IFS=$'\n'


case "$db_name" in
	"BKP.DISCOVERY")
		Command=$(ls -1 * |grep -vE "^$|:")
		if [ $? -eq 1 ]; then
			echo "Some error occured while list folder content"
			exit;
		fi
		;;
	*)
		Command=$(ls -l * |grep $db_name.sql.gz |awk '{print $9 " " $5}' )
		if [ $? -eq 1 ]; then
			echo "Some error occured while list folder content"
			exit;
		fi
		;;
esac

Command=$(echo $Command |sed -e 's/ /\n/g')
myJsonStr="{\"data\": ["

for listOutput in $Command
do
	dbName=$(echo $listOutput |awk '{print $1}'  | sed -e "s/.sql.gz//g" )
	dbSize=$(echo $listOutput |awk '{print $2}')

	JsonStr=$(echo $JsonStr\{ )
	JsonStr=$(echo $JsonStr\"{#DBNAME}\":\"$dbName\")
	if [ ! -z $dbSize ]; then
		JsonStr=$(echo $JsonStr\"{#DBSIZE}\":\"$dbSize\" )
	fi
	JsonStr=$(echo $JsonStr\}, )
done

myJsonStr=$(echo $myJsonStr $JsonStr |sed -e "s/,$//g")
echo $myJsonStr "] }"

IFS=$IFSBKP
