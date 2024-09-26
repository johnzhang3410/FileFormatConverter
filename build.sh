#automates the process of creating the build folder and running cmake
#!/bin/bash

if [ ! -d "build" ]; then
    mkdir build
fi

cd build

cmake ..
make
