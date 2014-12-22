#!/bin/bash
#
# This script get all Mysql database size (except the internal one) and display it as a Json
# that can be used to Zabbix Discovery Rules.
#
# This simple script was created by RafaelFoster: rafaelgfoster (at) gmail (dot) com
#
# Json Output: {"data": [ { "{#DBNAME}" : "database_name","{#DBSIZE}":"size_in_bytes"},{"{#DBNAME}":"database_name2","{#DBSIZE}":"size_in_bytes"} ] }

cd $HOME

SQL_CMD="AS 'Database', SUM(data_length + index_length) AS 'Size' FROM information_schema.TABLES"
SQL_ExceptionDB="information_schema|performance_schema|mysql"
db_name=$1

if [ $# -eq 0 ]; then
	db_name="DB.DISCOVERY"
fi

IFSBKP=$IFS
IFS=$'\n'

myJsonStr="{\"data\": ["

case "$db_name" in
	"DB.SIZE.TOTAL")
		SQL_Command="Select 'total' $SQL_CMD"
                echo $(mysql --skip-column-names -Ne  "$SQL_Command" | grep -vwE "$SQL_ExceptionDB" |awk '{print $2}' )
                exit;
		;;
	"DB.DISCOVERY")
		SQL_Command="Select table_schema FROM information_schema.TABLES GROUP BY table_schema"
		;;
	*)
		SQL_Command="Select table_schema $SQL_CMD WHERE table_schema = '$db_name' GROUP BY table_schema"
		echo $(mysql --skip-column-names -Ne  "$SQL_Command" | grep -vwE "$SQL_ExceptionDB" |awk '{print $2}' )
		exit;
		;;
esac

MYOUT=$(mysql -Ne "$SQL_Command" | grep -vwE "$SQL_ExceptionDB")
if [ $? -eq 1 ]; then
	echo "Some error occurred with mysql command"
	exit;
fi
for sqlOutput in $MYOUT
do
	dbName=$(echo $sqlOutput |awk '{print $1}')
	dbSize=$(echo $sqlOutput |awk '{print $2}')

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
