"""
Parse ASN csv file, and place them to the appropriate place in the Moterh File.
"""
import csv


class asnProduct():
    """
    Store one product from an asn.
    """
    def __init__(self, carton=None, pallet_no=None, color=None,
                 size=None, upc=None, coo=None, units=None, packs=None,
                 customer_po=None):
        self.carton = carton
        self.pallet_no = pallet_no
        self.color = color
        self.size = size
        self.upc = upc
        self.coo = coo
        self.units = units
        self.packs = packs
        self.customer_po = customer_po


    def __str__(self):
        return ('carton: {carton}, pallet_no: {pallet_no}, color: {color}, '
                'size: {size}, upc: {upc}, coo: {coo}, units: {units}, packs:'
                '{packs}, customer_po: {customer_po}'
                '').format(**self.__dict__,)

    def __repr__(self):
        return ('{__class__.__name__}(carton={carton!r}, pallet_no={pallet_no!r},'
                ' color={color!r}, size={size!r}, upc={upc!r}, coo={coo!r}, '
                'units={units!r}, packs={packs!r}, customer_po={customer_po!r})'
                '').format(__class__=self.__class__, **self.__dict__,)

    def parse_csv_row(self,row):
        self.carton = row[0]
        self.pallet_no = row[1]
        self.color = row[2]
        self.size = row[3]
        self.upc = row[4]
        self.coo = row[5]
        self.units = row[6]
        self.packs = row[7]
        self.customer_po = row[8]
        return self

class asnPackage():
    """
    store a gropu of items, that are under the same shipping reference number.
    """
    states = ['H1', 'H2', 'H3', 'ASNPROD'] 
    def __init__(self, shipping_ref=None, package_date=None, customer=None,
                 carrier=None, total_cartons=None, total_units=None,
                 total_packs=None, total_weight=None, sales_order=None,
                 product_list=None):
       self.shipping_ref = shipping_ref
       self.package_date = package_date
       self.customer = customer
       self.carrier = carrier
       self.total_cartons = total_cartons
       self.total_units = total_units
       self.total_packs = total_packs
       self.total_weight = total_weight
       self.sales_order = sales_order
       self.product_list = product_list if product_list else []

       self._state = 'H1'

    def __str__(self):
        tmp_str = '\n    '.join([str(prod) for prod in self.product_list])
        return ('shipping_ref: {shipping_ref}, package_date: {package_date}, '
                'customer: {customer}, carrier: {carrier}, total_cartons: '
                '{total_cartons}, total_units: {total_units}, total_packs: '
                '{total_packs}, total_weight: {total_weight}, sales_order: '
                '{sales_order} \nProduct_list:\n    {prod_list}'
                '').format(**self.__dict__, prod_list=tmp_str)
    
    def __repr__(self):
        return ('{__class__.__name__}('
                'shipping_ref={shipping_ref!r}, package_date={package_date!r}, '
                'customer={customer!r}, carrier={carrier!r}, total_cartons='
                '{total_cartons!r}, total_units={total_units!r}, total_packs='
                '{total_packs!r}, total_weight={total_weight!r}, sales_order='
                '{sales_order!r},product_list={product_list!r}'
                '').format(__class__=self.__class__, **self.__dict__,)

    def parse_csv_row(self,row):
        #print(row)
        if self._state == 'H1':
            #print('MIAFASZ')
            self.shipping_ref = row[1]
            self.package_date = row[2]
            self.customer = row[3]
            self.carrier = row[4]
            self.total_cartons = row[5]
            self.total_units = row[6]
            self.total_weight = row[7]
            self.total_packs = row[8]
            self._state = 'H2'
            return
        
        elif self._state == 'H2':
            self.sales_order = row[1]
            self._state = 'H3'
            return
        
        elif self._state == 'H3':
            self._state = 'ASNPROD'
            return
        
        elif self._state == 'ASNPROD':
            self.product_list.append(asnProduct().parse_csv_row(row))
            return

class asnList():
    """
    Parse/store the whole asn.pdf
    """

    def __init__(self, path=None, asn_package_list=None):
        """
        .
        """
        self.path = path
        self.asn_package_list = (asn_package_list if asn_package_list else [])
        self.parse_asn_file()
        
    def parse_asn_file(self):
        """
        return builds up
        """
        asn_pkg_elem = None
        with open(self.path) as csvfile:
            asnreader = csv.reader(csvfile, delimiter=',')
            for row in asnreader:
                if row[0] == 'HEADER1':
                    if asn_pkg_elem:
                        #print(asn_pkg_elem)
                        self.asn_package_list.append(asn_pkg_elem)
                    asn_pkg_elem = asnPackage()
                asn_pkg_elem.parse_csv_row(row)
            self.asn_package_list.append(asn_pkg_elem)
                
class preParser():
    """
    Send the appropriate lines from the csv to the appropriate parser.
    """
    def __init__():
        pass

