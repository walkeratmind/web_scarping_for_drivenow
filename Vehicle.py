
class Ride(object):
    def __init__(self, name="", supplier="", vehicle_type="",price=""):
        self.name = name
        self.supplier = supplier
        self.vehicle_type = vehicle_type
        self.price =price

    def get_name(self):
        return self.name
    
    def get_supplier(self):
        return self.supplier

    def get_vehicle_type(self):
        return self.vehicle_type

    def get_price(self):
        return self.price

    