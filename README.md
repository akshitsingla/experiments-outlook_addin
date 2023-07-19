# Build docker image
```sh
docker build -t helloworld_outlook_addin .
```

# Run docker container
```sh
docker run -v ./app:/app -w /app -it -p 3000:3000 --name helloworld_outlook_addin_container helloworld_outlook_addin
```

# Enter docker container
```sh
docker exec -it helloworld_outlook_addin_container sh 
```