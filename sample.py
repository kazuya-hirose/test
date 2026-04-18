# 本項で使用するコード

def calculate_sum(a, b):

    result = a + b

    return result

 

def main():

    x = 5

    y = 10

    z = calculate_sum(x, y)

 

    numbers = [1, 2, 3, 4, 5]

    numbers_squared = [n ** 2 for n in numbers]

 

    print(f"xとyの合計: {z}")

    print(f"二乗した数列: {numbers_squared}")

if __name__ == "__main__":

    main()