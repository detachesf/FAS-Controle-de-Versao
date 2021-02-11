import datapi as dap
import os, copy
import sagist as sg

DAT_KEYS = {}
DAT_KEYS["pds"]=DAT_KEYS["pdf"]=DAT_KEYS["pdd"]=DAT_KEYS["pas"]=DAT_KEYS["noh"]=DAT_KEYS["cnf"]=DAT_KEYS["nv1"]= \
    DAT_KEYS["tac"]=DAT_KEYS["paf"]=DAT_KEYS["cgs"]=DAT_KEYS["cgf"]=DAT_KEYS["pts"]=DAT_KEYS["ptf"]=\
    DAT_KEYS["pad"]=DAT_KEYS["lsc"]=DAT_KEYS["nv1"]=DAT_KEYS["nv2"]=DAT_KEYS["tdd"]=\
    DAT_KEYS["ctx"]=DAT_KEYS["enm"]=DAT_KEYS["grupo"]=DAT_KEYS["gsd"]=DAT_KEYS["ptd"]=DAT_KEYS["gsd"]= \
    DAT_KEYS["pro"]=DAT_KEYS["sev"]=DAT_KEYS["utr"]=DAT_KEYS["ins"]=DAT_KEYS["cxu"]=DAT_KEYS["enu"]= \
    DAT_KEYS["mul"]=DAT_KEYS["tn2"]=DAT_KEYS["tn1"]=DAT_KEYS["tcv"]=DAT_KEYS["ttp"]=DAT_KEYS["map"]= \
    DAT_KEYS["tctl"]=DAT_KEYS["tcl"]=['ID']
DAT_KEYS["rfi"]=DAT_KEYS["e2m"]=DAT_KEYS["grcmp"]=DAT_KEYS['ocr']=DAT_KEYS["sxp"]=DAT_KEYS["inp"]=DAT_KEYS["inm"]= \
    DAT_KEYS["psv"]=DAT_KEYS["cxp"]=[]
DAT_KEYS["rca"]=DAT_KEYS["rfc"]=['ORDEM','PNT','TPPNT']


def capitalize_keys(d):
    if type(d) == dict:
        for k in list(d.keys()):
            d[k.upper()] = d.pop(k) # coloca todas as chaves em maiúsculas
        return d
    else:
        raise TypeError('Dict expected')

def extract_keys(keys, data):
    keys_fields = {}
    for k in keys:
        if type(data) == DATPoint:
            field_val = data.get_value(k)
        elif type(data) == dict:
            field_val = data.get(k)
        if field_val is not None:
            keys_fields[k] = field_val
    if keys_fields == {}:
        return None
    else:
        return keys_fields

def check_keys(dats, data):
    '''
    Retorna False se o objeto data possui campos chaves existentes no objeto dats
    :param dats: objeto tipo DAT ou DATCollection
    :param data: objeto tipo Point ou dict
    :return: bool = False se as chaves existem, True se não existe 
    '''
    if type(dats) not in [DATCollection, DAT]:
        raise TypeError('DAT or DATCollection expected')
    datcol = type(dats) == DATCollection
    if datcol:
        keys = dats.root.keys
    else:
        keys = dats.keys
    keys_fields = extract_keys(keys, data)
    return (keys_fields is None) or (dats.find_all(keys_fields) is None)

def validate_point_input(data):
    '''
    Trata um valor de entrada para checar se é um valor de ponto válido para ser tratado por outras funções.
    :param data: str, dict ou Point 
    :return: dict ou Point
    '''
    if type(data)==str:
        data = {'ID':data}

    if type(data) not in [str, DATPoint, dict]:
        raise TypeError('Point must be str, dict or DATPoint type')

    if type(data) == dict:
        data = capitalize_keys(data)

    return data


class DatField():
    def __init__(self, name, value=''):
        self.name = name.upper()
        self.value = value
        self.__enable = True
        self.comment = ''

    @property
    def enable(self):
        return self.__enable

    @enable.setter
    def enable(self, value):
        if type(value) == bool:
            self.__enable = value
        else:
            raise TypeError('Boolean expected')

    def __str__(self):
        return self.as_str()

    def print(self):
        print(self.as_str())

    def as_str(self, comment=True):
        r = '{} = {}'.format(self.name, self.value)
        if not self.enable:
            r = ';'+r
        if (comment) and (self.comment !=''):
            r = ';' + self.comment + '\n' + r
        return r




class Point():
    def __init__(self, data=None, parent=None):
        #self.__type = dattype.upper()
        self.__fields = []
        self.__field_count = 0
        self.parent = parent
        self.__enable = True
        self.comment = ''

        if type(data) == dict:
            self.add_fields(data)
        elif type(data) == str:
            self.add_field(name='id', value=data)

    @property
    def enable(self):
        return self.__enable


    def field(self, name):
        i = self.indexof(name)
        if i is not None:
            return self[i]
        else:
            raise LookupError('Field does not exist')

    @enable.setter
    def enable(self, value):
        if type(value) == bool:
            self.__enable = value
            for f in self:
                f.enable = value
        else:
            raise TypeError('Boolean expected')


    @property
    def field_count(self):
        return self.__field_count

    #@property
    #def type(self):
    #    return self.__type



    def __getitem__(self, item):
        if isinstance(item, int):
            return self.__fields[item]
        elif isinstance(item, str):
            return self.get_value(item)

    def __setitem__(self, key, value):
        if isinstance(key, int):
            self.__fields[key] = value
        elif isinstance(key, str):
            self.add_or_update(key,value)

    def fieldnames(self):
        l = [p.name for p in self.__fields]
        return l

    def field_exists(self,name):
        name = name.upper()
        return name in self.fieldnames()

    def add_field(self, name, value):
        name = name.upper()
        if self.field_exists(name):
            raise
        else:
            field = DatField(name, value)
            self.__fields.append(field)
            self.__field_count +=1

    def add_fields(self, fields):
        if type(fields) == dict:
            for key in fields.keys():
                self.add_field(name=key, value=fields[key])

    def indexof(self, field):
        field = field.upper()
        i=0
        for p in self.__fields:
            if p.name == field:
                return i
            i+=1
        return None


    def delete_field(self, field):
        i = self.indexof(field)
        if i is not None:
            self.__fields.pop(i)
            self.__field_count-=1

    def get_value(self, field):
        field = field.upper()
        i = self.indexof(field)
        if i is not None:
            return self.__fields[i].value
        else:
            return None

    def set_value(self, field, value):
        field = field.upper()
        i = self.indexof(field)
        if i is not None:
            self.__fields[i].value = value
        else:
            raise KeyError('Field does not exist')

    def add_or_update(self, field, value):
        if self.field_exists(field):
            self.set_value(field, value)
        else:
            self.add_field(field, value)

    def print(self, comment=True):
        print(self.as_str(comment=comment))

    def as_str(self, comment=True):
        #r = self.type+'\n'
        r=''
        for f in self.__fields:
            r = r+f.as_str()+'\n'

        if not self.enable:
            r = ';'+r
        if (comment) and (self.comment != ''):
            r = ';' + self.comment + '\n'+r

        return r

    def as_dict(self, do_comment=True):
        r = {}
        for f in self:
            name = f.name
            if f.enable:
                r[name] = f.value
            elif do_comment:
                name = ';'+name
                r[name]=f.value
        return r

    def remove_blank_fields(self):
        for f in self.fieldnames():
            if self.get_value(f)=='':
                self.delete_field(f)

    def remove_disabled_fields(self):
        names = self.fieldnames()
        for f in names:
            if self[self.indexof(f)].enable == False:
                self.delete_field(f)

    def clear(self):
        self.remove_blank_fields()
        self.remove_disabled_fields()


    def __eq__(self, other):

        isequal = True

        this = self.copy()
        that = other.copy()
        this.clear()
        that.clear()

        # devem ter os mesmos campos válidos
        isequal = isequal and (set(this.fieldnames()) == set(that.fieldnames()))

        # devem ter o mesmo tipo
        isequal = isequal and (this.type == that.type)

        # devem estar ativos/inativos
        isequal = isequal and (this.enable == that.enable)

        for field in this.fieldnames():
            isequal = isequal and (this.get_value(field) == that.get_value(field))

        return isequal

    def fields_match(self, fields, tokens=''):
        '''
        Retorna True se os campos passados como parâmetro batem com o do ponto.
        :param fields: dict cujas chaves são os campos em que se quer testar, com respectivos valores de procura
        :param tokens: 'l' (like) faz com que campos que contenham os valores passados como parâmetro retornem true,
        mesmo que não seja exatamente igual
        :return: 
        '''
        isequal = True
        for f in list(fields.keys()):
            if 'L' in tokens.upper():
                isequal = isequal and (fields[f] in self.get_value(f))
            else:
                isequal = isequal and (fields[f] == self.get_value(f))
        return isequal


    def __str__(self):
        return self.as_str()

    def copy(self):
        return copy.deepcopy(self)

    def clone(self, replace=None, fields=None):
        new = self.copy()
        new.replace(values=replace, fields=fields)
        self.parent.add_point(new)


    def replace(self, values=None, fields=None):
        '''
        Substitui os valores dos campos de um ponto. Os campos que devem sofrer alteração são
        passados como uma lista, e os valores a serem subsituídos são passados como um par (antigo, novo)
        :param values: list ou tuple com os valores que serão substituídos, no formato (antigo, novo)
        :param fields: list ou tuple com os campos que sofrerão a substituição ou 'all' para todos os campos. Se
        deixado em branco, apenas a chave da tabela é alterada
        :return: A substituição é feita no próprio objeto
        '''
        if fields is None:
            if 'ID' in self.fieldnames():
                fields = ['ID']
            else:
                fields = self.fieldnames()
        elif fields == 'all':
            fields = self.fieldnames()
        elif type(fields) not in (list, tuple):
            raise TypeError('List or tuple expected')
        for field in fields:
            v = self.get_value(field)
            if v is not None:
                if values is None:
                    self.set_value(field, v + '_CLONE')
                elif (type(values) in (list, tuple)) and len(values) == 2:
                    self.set_value(field, v.replace(values[0], values[1]))



class DATPoint(Point):
    def __init__(self, dattype, data=None, parent=None):
        super().__init__(data=data, parent=parent)
        self.__type = dattype.lower()
        names = self.fieldnames()
        for f in dap.DAT_FIELDS[dattype.lower()]:
            if f not in names:
                self.add_field(name=f, value='')

    @property
    def type(self):
        return self.__type

    def print(self, comment=True):
        tstr = self.type.upper()+'\n'
        pstr = self.as_str(comment=comment)
        print(tstr+pstr)

class DAT():
    def __init__(self, dattype, dir='', name='', parent=None):
        self.__type = dattype.lower()
        self.__points = []
        self.parent = parent
        self.dir = dap.fix_path(dir)
        if name == '':
            name = self.dattype + '.dat'
        self.name = name
        self.includes = []
        self.keys = DAT_KEYS[self.__type]


    #def points(self, search):
    #    result = []


    @property
    def path(self):
        return os.path.join(self.dir, self.name)

    @property
    def dattype(self):
        return self.__type


    @property
    def total_points(self):
        return len(self.__points)

    @property
    def total_enabled(self):
        i=0
        for p in self:
            if p.enable:
                i +=1
        return i


    def point(self, search):
        i = self.indexof(search)
        if i is not None:
            return self[i]
        else:
            return None

    def as_dict(self, do_comment=True):
        return [p.as_dict(do_comment=do_comment) for p in self]


    def indexof(self, key):
        key = validate_point_input(key)

        if type(key) == dict:
            i = 0
            for p in self:
                isequal = True
                for k in list(key.keys()):
                    isequal = (isequal) and (key[k] == p.get_value(k))
                if isequal:
                    return i
                i+=1
        elif type(key) == DATPoint:
            i = 0
            for p in self:
                if p.fields_match(key.as_dict()):
                    return i
                i += 1

        return None

    # função antiga
    def __indexof_old(self, key):
        key = validate_point_input(key)

        if type(key) == dict:
            i = 0
            for p in self:
                isequal = True
                for k in list(key.keys()):
                    isequal = (isequal) and (key[k] == p.get_value(k))
                if isequal:
                    return i
                i+=1
        elif type(key) == DATPoint:
            i = 0
            for p in self:
                if p == key:
                    return i
                i += 1

        return None

    def find_all(self, data, tokens=''):
        data = validate_point_input(data)
        result = []
        for point in self:
            if type(data) == dict:
                found = point.fields_match(data, tokens=tokens)
            else:
                found = data == point
            if found:
                result.append(point)
        if result:
            return result
        else:
            return None


    def delete_point(self, key):
        if type(key) == int:
            i = key
        else:
            i = self.indexof(key)
        if i is not None:
            self.__points.pop(i)


    def delete_pointset(self, key):
        if type(key) not in [list, tuple]:
            key = [key]
        for p in key:
            self.delete_point(p)

    def delete_where(self, key, tokens=''):
        s = self.find_all(data=key, tokens=tokens)
        self.delete_pointset(s)





    def add_point(self, point, check=True):

        point = validate_point_input(point)

        if (type(point) == dict) or (type(point) == str):
            point = DATPoint(dattype=self.__type, data=point, parent=self)

        #if type(point) == DATPoint:
        #    if self.keys != ['']:
        #        p = {}
        #        for k in self.keys:
        #            p[k] = point.get_value(k)
        #        i = self.indexof(p)
        #        if i is not None:
        #            raise KeyError('DAT keys already in use')
        #
        #    self.__points.append(point)
        #else:
        #    raise TypeError('DATPoint expected')

        if self.parent is not None:
            searchobj = self.parent
        else:
            searchobj = self

        if check and not check_keys(searchobj, point):
            raise TypeError('DAT keys {} already in use'.format(point.as_str()))
        else:
            point.parent = self
            self.__points.append(point)



    def __str__(self):
        return self.as_str()

    def __getitem__(self, item):
        return self.__points[item]

    def __setitem__(self, key, value):
        self.__points[key]=value


    def as_str(self, comment=True):
        r = ''
        for p in self:
            r = r + p.as_str(comment=comment)
            r = r + '\n'
        return r


    def print(self, comment=True):
        print(self.as_str(comment=comment))


    def remove_disabled(self):
        erase = []
        #i = 0
        for p in self:
            if p.enable == False:
                erase.append(p.copy())
        #    i += 1

        for p in erase:
            self.delete_point(p)

    def remove_blanks(self):
        for p in self:
            p.remove_blank_fields()

    def copy(self):
        return copy.deepcopy(self)

    def clear(self):
        self.remove_blanks()
        self.remove_disabled()

    def delete_all(self):
        self.__points = []


    def make_iccpid(self):
        if self.dattype.upper() not in ['PDS','PTS','PAS','CGS']:
            raise TypeError('DAT must be pds, pas, cgs or pts')
        if self.dattype.upper() == 'CGS':
            sufix = 'C'
        else:
            sufix = ''

        for p in self:
            p.add_or_update('ICCPID', p.get_value('ID').replace(':','_')+sufix)



    def __eq__(self, other):
        isequal = True

        #isequal = isequal and (self.path == other.path)
        isequal = isequal and (self.dattype == other.dattype)

        this = self.copy()
        that = other.copy()

        this.clear()
        that.clear()

        isequal = isequal and (this.total_points == that.total_points)

        for p in this:
            isequal = isequal and (p in that)

        return isequal


    def clone(self, dir, name=None, replace=None, fields=None):
        new_dat = self.copy()
        dir = dap.fix_path(dir)
        new_dat.dir = dir
        if name is None:
            name = self.dattype.lower()+'.dat'
        elif not isinstance(name, str):
            raise TypeError('Str expected')
        new_dat.name = name
        for p in new_dat:
            p.parent = new_dat
            p.replace(values=replace, fields=fields)
        self.parent.add_dat(new_dat)



class DATCollection():
    def __init__(self, dattype, dats=[]):
        self.__dats = []
        if dats:
            self.__dats = dats.copy()
            for dat in dats:
                dat.parent = self
        self.dattype = dattype.lower()

    def indexof(self, key):
        key = dap.fix_path(key)
        name = os.path.basename(key)
        dir = os.path.dirname(key)
        c = 0
        for item in self.__dats:
            if (item.name == name) and (item.dir == dir):
                return c
            c +=1
        return None

    def __getitem__(self, item):
        if isinstance(item, int):
            return self.__dats[item]
        elif isinstance(item, str):
            for path in self.list_paths():
                if item in path:
                    return self.dat(path)
        else:
            raise TypeError('Str or int expected')
        return None



    def __setitem__(self, key, value):
        if type(value) != DAT:
            raise TypeError('DAT expected')
        i = self.indexof(value.path)
        if i is not None:
            if i != key:
                raise ValueError('DAT path {} already in use'.format(value.path))
        self.__dats[key] = value

    def __len__(self):
        return self.total_dats


    @property
    def root(self):
        return self.dat()

    @property
    def all(self):
        l = []
        for dat in self:
            for p in dat:
                l.append(p)
        return l

    def dat(self, key=''):
        if not key:
            key = '{}.dat'.format(self.dattype)
        return self.__dats[self.indexof(key)]

    @property
    def total_dats(self):
        return len(self.__dats)

    def remove_blanks(self):
        for dat in self:
            dat.remove_blanks()

    def remove_disabled(self):
        for dat in self:
            dat.remove_disabled()

    def clear(self):
        self.remove_disabled()
        self.remove_blanks()


    def point(self, key):
        for dat in self:
            p = dat.point(key)
            if p:
                return p
        return None

    def find_all(self, data, tokens=''):
        l = []
        for dat in self:
            r = dat.find_all(data, tokens=tokens)
            if r is not None:
                l.extend(r)
        if l:
            return l
        else:
            return None

    def delete_point(self, key):
        for dat in self:
            dat.delete_point(key)


    def delete_pointset(self, key):
        for dat in self:
            dat.delete_pointset(key)


    def add_point(self, data, dat=None):
        if dat is None:
            dat = self.dattype + '.dat'
        if type(dat) != str:
            raise TypeError('Str expected')

        data = validate_point_input(data)

        # checa se já existe ponto entre os dats com mesmas chaves

        if not check_keys(self, data):
            raise KeyError('DAT keys already in use')
        else:
            self.dat(dat).add_point(data)


    def list_paths(self):
        return [dat.path for dat in self]


    def has_include(self, include):
        r = False
        for path in self.list_paths():
            r = r or include in path
        return r

    def add_include(self, include):
        if self.has_include(include):
            raise KeyError('Include already exists in dat collection')
        new_dat = DAT(dattype=self.dattype, dir=include)
        self.add_dat(new_dat)

    def add_dat(self, dat):
        if dat.path in self.list_paths():
            raise ValueError('DAT path {} already in use'.format(dat.path))
        if type(dat) != DAT:
            raise TypeError('DAT type expected')
        self.__dats.append(dat)
        dat.parent = self

    def delete_dat(self, key):
        if type(key) == int:
            i = key
        else:
            i = self.indexof(key)
        if i is not None:
            self.__dats.pop(i)

    def as_str(self, comment=True):
        r = ''
        for dat in self:
            r += '# {} \n\n'.format(dat.path)
            r += dat.as_str(comment=comment)
        return r

    def as_dict(self, do_comment = True):
        r = {}

        for dat in self:
            r[dat.path] = dat.as_dict(do_comment=do_comment)
        return r


    def print(self, comment=True):
        print(self.as_str(comment=comment))

    def make_iccpid(self):
        if self.dattype.upper() not in ['PDS','PAS','PTS','CGS']:
            raise TypeError('DAT collection must be of pds, pas, pts or cgs type')
        for dat in self:
            dat.make_iccpid()


    def __str__(self):
        return self.as_str()

    @property
    def total_points(self):
        total = 0
        for dat in self:
            total += dat.total_points
        return total

    @property
    def total_enabled(self):
        total = 0
        for dat in self:
            total += dat.total_enabled
        return total

    def delete_all(self):
        self.__dats = []
        self.__dats.append(DAT(self.dattype))

    def write_dats(self, path, verbose=False):
        d = self.as_dict()
        for k in list(d.keys()):
            if k == '{}.dat'.format(self.dattype):
                d[self.dattype] = d.pop(k)
            else:
                d['#{}'.format(k)] = d.pop(k)

        dap.write_dat(self.dattype, d, source_path=path, dests=sorted(list(d.keys())), verbose=verbose, do_backup=False)





class Base():

    def __init__(self, name='demo', path='.', populate=False, source=None, check=True):
        self.name = name
        self.path = path

        self.__dat_list = list(DAT_KEYS.keys())

        #self.dats_path = 'demo'

        for dat_name in self.__dat_list:
            setattr(self, dat_name, DATCollection(dat_name,[DAT(dat_name)]))
        if populate:
            self.populate_demo()
        if source is not None:
            self.read_dats(source=source, check=check)


    @property
    def dat_list(self):
        return self.__dat_list.copy()

    def populate_demo(self):
        l = []
        for i in range(10):
            p = {
                'id':'id{}'.format(i),
                'nome': 'nome{}'.format(i),
                'ocr': 'ocr1',
                'tac': 'tac1',
                'cmt': '{}'.format(i)
            }
            l.append(p)
        for p in l:
            self.pds.root.add_point(p)
            self.pas.root.add_point(p)


    @property
    def path(self):
        return self.__path


    @path.setter
    def path(self, value):
        if not isinstance(value, str):
            raise TypeError('Str expected')
        value = dap.fix_path(value)
        if value == '':
            value = '.'
        self.__path = value


    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, value):
        if not isinstance(value, str):
            raise TypeError('Str expected')
        if value == '':
            value = 'demo'
        self.__name = value


    @property
    def total_points(self):
        total = 0
        for dat in self.__dat_list:
            datcoll = getattr(self, dat)
            total += datcoll.total_points
        return total

    @property
    def total_enabled(self):
        total = 0
        for dat in self.__dat_list:
            datcoll = getattr(self, dat)
            total += datcoll.total_enabled
        return total

    def delete_all(self):
        for dat in self.__dat_list:
            datcoll = getattr(self, dat)
            datcoll.delete_all()

    def list_includes(self):
        l = []
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            for path in datcoll.list_paths():
                p = os.path.dirname(path)
                if p not in ['','.']:
                    l.append(p)
        l = list(set(l))
        if l:
            return l
        else:
            return None

    def del_include(self, key):
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            path_list = datcoll.list_paths()
            for path in path_list:
                if key in os.path.dirname(path):
                    datcoll.delete_dat(path)




    def read_dats(self, source, check=True):
        self.path = source
        b = sg.load_base(source_path=source, no_comments=True, verbose=False)
        self.delete_all()
        for datname in list(b.keys()):
            datcoll = getattr(self, datname)
            for key in list(b[datname]):
                path = key
                if not '.dat' in path:
                    path = path + '.dat'
                path = path.lstrip('#')
                path = dap.fix_path(path)
                name = os.path.basename(path)
                dir = os.path.dirname(path)
                datobj = DAT(datname, dir=dir, name=name)
                for p in b[datname][key]:
                    #print(p)
                    datobj.add_point(p, check=check)
                if path == datcoll.root.path:
                    datcoll[0] = datobj.copy()
                else:
                    datcoll.add_dat(datobj.copy())

    def write_dats(self, path='', verbose=False):
        if path == '':
            path = self.path

        #path = os.path.join(self.path, self.name)

        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            datcoll.write_dats(path, verbose=verbose)


    def remove_blanks(self):
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            datcoll.remove_blanks()

    def remove_disabled(self):
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            datcoll.remove_disabled()

    def clear(self):
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            datcoll.clear()


    def make_iccpid(self):
        for datname in self.dat_list:
            if datname.upper() in ['PDS','PAS','PTS','CGS']:
                datcoll = getattr(self, datname)
                datcoll.make_iccpid()


    def clone_include(self, source, to, replace=None, fields=None):
        for include in self.list_includes():
            if to in include:
                raise KeyError('Include dir already exists')
        for datname in self.dat_list:
            datcoll = getattr(self, datname)
            for dat in datcoll:
                if source in dat.dir:
                    dat.clone(dir=to, replace=replace, fields=fields)




    def add_dnp3_d(self, name, include=None):

        gsd_id = self.gsd.root[0]['ID']
        lsca = {}
        lscd = {}
        lsca['gsd']=lscd['gsd']=gsd_id
        lsca['id'] = name+'_A'
        lscd['id'] = name+'_D'
        lsca['nome'] = 'Ligacao de aquisicao virtual {}'.format(name)
        lscd['nome'] = 'Ligacao de distribuicao {}'.format(name)
        lsca['nsrv1'] = lsca['nsrv2'] = lscd['nsrv1'] = lscd['nsrv2'] = 'localhost'
        lsca['tcv'] = lscd['tcv'] = 'CNVH'
        lsca['tipo'] = 'AA'
        lscd['tipo'] = 'DD'
        lsca['ttp'] = 'CXTCP'
        lscd['ttp'] = 'IEC3S'
        lsca['verbd'] = lscd['verbd'] = 'SAGIST'
        lsca['map'] = lscd['map'] = 'GERAL'

        confa = self.get_new_serial_lines()

        cnfa = {}
        cnfd = {}

        cnfa['config']= 'PlPr= {} LiPr= {} PlRe= {} LiRe={} TZBR= 0 DnpLvl= 3'.format(cnfa['PlPr'], cnfa['LiPr'], cnfa['PlRe'], cnfa['LiRe'])
        cnfd['config']= 'PlPr= {} LiPr= {} PlRe= {} LiRe={} TZBR= 0 DnpLvl= 3'.format(cnfa['PlPr'], cnfa['LiPr']+1, cnfa['PlRe'], cnfa['LiRe']+1)
        cnfa['id'] = name+'_A'
        cnfd['id'] = name+'_D'
        cnfa['lsc'] = lsca['id']
        cnfd['lsc'] = lscd['id']
        

        if include is None:
            self.lsc.root.add_point(lsca)
            self.lsc.root.add_point(lscd)
        else:
            if not self.lsc.has_include(include):
                self.lsc.add_include(include)
            self.lsc[include].add_point(lsca)
            self.lsc[include].add_point(lscd)


    def get_serial_lines(self):
        lp = []
        lr = []
        for p in self.cnf.all:
            if 'PlPr=' in p['config']:
                l = p['config'].split(' ')
                try:
                    plpr = int(l[l.index('PlPr=')+1])
                except:
                    plpr = None
                try:
                    lipr = int(l[l.index('LiPr=')+1])
                except:
                    lipr = None
                try:
                    plre = int(l[l.index('PlRe=')+1])
                except:
                    plre = None
                try:
                    lire = int(l[l.index('LiRe=')+1])
                except:
                    lire = None
                lp.append((plpr,lipr))
                lr.append((plre,lire))
        d = {}
        d['P'] = lp
        d['R'] = lr
        return d

    def show_serial_lines(self):
        for dat in self.cnf:
            include = dat.path
            first = True
            for p in dat:
                config = p['config']
                if 'PlPr=' in config:
                    if first:
                        print('\n'+include)
                        first = False
                    print(p['id']+': '+p['config'])


    def get_new_serial_lines(self, as_string=False):
        d = self.get_serial_lines()
        plpr = [p[0] for p in d['P']]
        plre = [p[0] for p in d['R']]
        d={}
        plpr.sort()
        plre.sort()
        d['PlPr'] = plpr[-1]+1
        d['LiPr'] = 1
        d['PlRe'] = plpr[-1]+2
        d['LiRe'] = 1
        if as_string:
            s = ''
            for key in d.keys():
                s = s+str(key)+'= '+str(d[key])+' '
            return s.strip()
        else:
            return d


    def get_dist_points(self, tdd=None):
        if (type(tdd) != str) or (tdd==''):
            raise TypeError('Parâmetro tdd deve ser uma string válida')

        pdds = self.pdd.find_all({'TDD':tdd})
        pads = self.pad.find_all({'TDD': tdd})
        ptds = self.ptd.find_all({'TDD': tdd})

























