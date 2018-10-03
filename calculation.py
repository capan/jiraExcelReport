import datetime as d
import numpy as np

ymd_create = []
ymd_resdate = []
delta_t = []

class calculate:
    def __init__(self):
        self.result = 0

    def meantime(self, issueobject):
        for i in range(0, len(issueobject)):
            ymd_create.append(d.datetime(int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                     [u'created'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[1])))
            ymd_resdate.append(d.datetime(int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                                    [u'resolutiondate'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[1])))
            delta_t.append((ymd_resdate[i] - ymd_create[i]).days)

        self.result = np.mean(np.array(delta_t))
        return self.result

    def mediantime(self, issueobject):
        for i in range(0, len(issueobject)):
            ymd_create.append(d.datetime(int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                     [u'created'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[1])))
            ymd_resdate.append(d.datetime(int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                                    [u'resolutiondate'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[1])))
            delta_t.append((ymd_resdate[i] - ymd_create[i]).days)

        self.result = np.median(np.array(delta_t))
        return self.result

    def variancetime(self, issueobject):
        for i in range(0, len(issueobject)):
            ymd_create.append(d.datetime(int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                     [u'created'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'created'].split('T')[1].split(':')[1])))
            ymd_resdate.append(d.datetime(int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[0].split('-')[1]), int(issueobject[i].raw[u'fields']
                                                                                                                                                                                                                    [u'resolutiondate'].split('T')[0].split('-')[2]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[0]), int(issueobject[i].raw[u'fields'][u'resolutiondate'].split('T')[1].split(':')[1])))
            delta_t.append((ymd_resdate[i] - ymd_create[i]).days)
        
        self.result = np.var(np.array(delta_t))
        return self.result
