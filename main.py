#!/usr/bin/env python3
import os.path
import cherrypy
import json
import xlrd
from functools import partial
from jinja2 import Template, Environment, FileSystemLoader

class Root:
    @cherrypy.expose
    def index(self):
        return jj.get_template('index.html').render()

    @cherrypy.expose
    def search_school(self, keyword, scholarship, private_or_public, location):
        result = [s for s in schools if check(s, keyword, scholarship, private_or_public, location)]
        return jj.get_template('search_school.html').render(schools=result)

    @cherrypy.expose
    def search_med(self, keyword, degree, country, focus):
        focus = int(focus)
        data = xlrd.open_workbook(os.path.join(os.path.dirname(__file__), "med.xlsx"))
        table = data.sheets()[0]
        ans_list = []

        for i in range(table.nrows):
            if i > 0 and check_med(table, i, keyword, country, degree):
                ans_list.append(table.row(i))

        ans_list.sort(key=lambda med: -med[focus].value)

        meds = []
        for i in ans_list:
            med = []
            for j in range(table.ncols):
                a = table.cell(0, j)
                b = i[j]
                med.append((a.value, b.value))
            meds.append(med)

        return jj.get_template('search_med.html').render(meds=meds)

def findTheKeyWord(node, inp):
    if type(node) == dict:
        for item in node.values():
            if findTheKeyWord(item, inp):
                return True
    elif type(node) == list or type(node) == set:
        for item in node:
            if findTheKeyWord(item, inp):
                return True
    elif type(node) == str:
        if node.find(inp) != -1:
            return True

    return False

def check(node, inp, scholarship, private_or_public, location):
    if scholarship != '':
        for item in node['attrs']:
            if item.find('是否有奖学金') != -1 and item[6:].find(scholarship) == -1:
                return False
    if private_or_public != '':
        for item in node['attrs']:
            if item.find('院校性质') != -1 and item.find(private_or_public) == -1:
                return False
    if location != '':
        for item in node['attrs']:
            if item.find('地理位置') != -1 and item.find(location) == -1:
                return False
    return findTheKeyWord(node, inp)

def check_med(table, i, s, country, degree):
    if table.cell(i, 3).value.find(country) == -1:
        return False
    if table.cell(i, 4).value.find(degree) == -1:
        return False

    for j in range(table.ncols):
        a = table.cell(i, j)
        if a.ctype == 1:
            if a.value.find(s) != -1:
                return True

    return False

def load_schools():
    global schools
    file_object = open(os.path.join(os.path.dirname(__file__), "bailitop.json"))
    content = file_object.read()
    schools = json.loads(content)

    for school in schools:
        s = school['apply']
        p = s.find('托福最低分数')
        if p != -1:
            p += 7
            school['tofu'] = ''
            while (s[p] != '<'):
                school['tofu'] += s[p]
                #IPython.embed()
                p += 1

        p = s.find('雅思分数要求')
        if p != -1:
            p += 7
            school['yasi'] = ''
            while (s[p] != '<'):
                school['yasi'] += s[p]
                p += 1

load_schools()
jj = Environment(loader=FileSystemLoader(os.path.abspath(os.path.dirname(__file__))))
cherrypy.config.update({'engine.autoreload.on': False})
cherrypy.quickstart(Root(), '/', {
    '/': {
        'tools.sessions.on': True,
        'tools.staticdir.root': os.path.abspath(os.path.dirname(__file__))
    },
    '/css': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': 'css'
    },
    '/fonts': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': 'fonts'
    },
    '/js': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': 'js'
    },
})
