sudo docker container prune -fa
sudo docker images prune -fa
sudo docker network prune -fa
cd ~
sudo rm -r temp
git clone https://github.com/AndreasL-GL/temp.git
cd temp
sudo docker compose up --build