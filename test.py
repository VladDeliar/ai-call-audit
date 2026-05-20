import torch

print(torch.__version__)                  # має бути типу 2.x.x+cu128
print(torch.cuda.is_available())          # True
print(torch.cuda.get_device_name(0))      # NVIDIA GeForce GTX 1050 Ti
print(torch.version.cuda)                 # 12.8