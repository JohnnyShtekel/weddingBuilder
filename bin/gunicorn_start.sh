#!/bin/sh


NAME="gvia_yadim_report"
SERVERDIR=/home/tal/apis/gvia_yadim_report
NUM_WORKERS=3
USER=tal
GROUP=www-data

echo "Starting $NAME as `whoami`"
echo "`pwd`"

cd $SERVERDIR

source ../.././.virtualenvs/gvia_yadim_report/bin/activate
export PYTHONPATH=$SERVERDIR:$PYTHONPATH

exec ../.././.virtualenvs/gvia_yadim_report/bin/gunicorn wsgi:app --name $NAME --workers $NUM_WORKERS --user=$USER -b 0.0.0.0:8021 --log-level=debug --log-file=- --timeout=600
