sudo docker container prune
sudo docker images prune
sudo docker network prune
cd ..
sudo rm -r temp
git clone https://github.com/AndreasL-GL/temp.git
cd temp
sudo docker compose up -d --build