sudo docker container prune -y
sudo docker images prune -y
sudo docker network prune -y
cd ~
sudo rm -r temp
git clone https://github.com/AndreasL-GL/temp.git
cd temp
sudo docker compose up --build