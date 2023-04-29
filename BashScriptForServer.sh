cd ~
if [ ! -d ~/items/trädgårdsexperterna ]; then
  mkdir -p ~/items/trädgårdsexperterna
fi
container_id=$(sudo docker container ls -f name=api -q)
sudo docker cp $container_id:/usr/src/app/functions/Excel/items_trädgårdsexperterna ~/items/trädgårdsexperterna

sudo rm -r temp
git clone https://github.com/AndreasL-GL/temp.git
cd temp
cd api
sudo python3 app.py