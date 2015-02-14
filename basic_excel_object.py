class BasicObject:

    _reg_clsid_ = '{D67CBA29-8DE9-453E-A2FE-6926AC5D2B0E}'
    _reg_progid_ = "Python.Object"
    _public_methods_ = ["fibonacci",
                        "pylookup",
                        "pycount"]

    def __init__(self):
        self.count = 0

    def fibonacci(self, n):
        """Returns the nth number in the fibonacci sequence"""

        if n == 1:
            return 1
        elif n <= 0:
            return 0
        else:
            return self.fibonacci(n - 1) + self.fibonacci(n - 2)

    def pylookup(self, col1, col2, matrix, index=3):
        """Similiar to VLOOKUP except there are 2 key values opposed to 1"""

        for row in matrix:
            if col1 == row[0] and col2 == row[1]:
                return row[index]
        return None

    def pycount(self):
        """A simple method that increments self.count"""

        self.count += 1
        return self.count

if __name__ == "__main__":
    from win32com.server.register import UseCommandLine
    UseCommandLine(BasicObject)