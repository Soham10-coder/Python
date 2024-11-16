import matplotlib.pyplot as plt
plt.plot([1, 2, 3], [4, 5, 6], color="red")
plt.xlim(1, 5)
plt.ylim(1, 6)
plt.xlabel('X-axis')
plt.ylabel('Y-axis')
plt.title('Demo Graph')
plt.legend(['Values'])
plt.show()
