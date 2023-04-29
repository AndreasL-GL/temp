cd ~
if [ ! -d ~/items/trädgårdsexperterna ]; then
  mkdir -p ~/items/trädgårdsexperterna
fi
container_id=$(sudo docker container ls -f name=api -q)
sudo docker container cp -r $container_id:/usr/src/app/functions/Excel/items_trädexperterna/ ~/items/trädexperterna/.
sudo docker compose down
sudo rm -r temp
git clone https://github.com/AndreasL-GL/temp.git
cd temp
sudo docker compose up --build -d