set currentdate=%date:~0,10%
set commitmsg="update %currentdate%"
e:
cd e:\projects-home\reading
git pull
git add ./
git commit -am %commitmsg%
git push