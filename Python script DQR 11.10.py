# -*- coding: utf-8 -*-
"""
Created on Thu Apr  3 11:08:40 2025

@author: Derek - uqu2. Adapted for urb7, November 2025
"""


#import argparse
import json
import os
import datetime
import pandas as pd
#from tkinter import filedialog
from collections import Counter, defaultdict



'''
Parameters used when running the script.
'''

# Add argparse to handle command-line arguments
'''
parser = argparse.ArgumentParser(description='Process healthcare data bundles.')
parser.add_argument('--mode', choices=['single', 'batch'], default='single',
                    help='Specify processing mode: "single" for individual file or "batch" for multiple files in a directory.')
parser.add_argument('--facility', default='',
                    help='Specify facility name.')
parser.add_argument('--month_year', default='',
                    help='Specify 3 letter month _ 4 digit year.')
parser.add_argument('--Molly', choices=['False', 'True'], default='False', help='Use True if this is for Molly')
args = parser.parse_args()'''



'''
Function to unpack each line of the ndjson file and log the summary data into 
the respective dataframes.

input: line from ndjson file (as a json object)
output: a dictionary with keys for each unique element path found in the input line

Note: function is iterative, and will call itself as needed to fully unpack nested objects
'''

def unpack(data, element_path=None, results=None):
    if results is None:
        results = defaultdict(list)

    if element_path is None:
        element_path = ''

    if isinstance(data, dict):
        for key, value in data.items():
            if not isinstance(key, int):
                unpack(value, f"{element_path}_{key}" if element_path else key, results)
            else:
                unpack(value, element_path, results)
            

    elif isinstance(data, list):
        for item in data:
            unpack(item, element_path, results)

    else:
        results[element_path].append(data) 

    
    return dict(results)


'''
Function takes the results of the unpack() function to create counters to populate
the summary dataframe.

input: a dictionary
output: a dictionary of counters
'''
def get_Vals(z):
    for key in z['addl_info'].keys():
        if type(z['addl_info'][key][0]) != list and type(z['addl_info'][key][0]) != dict:
            z['vals'][key] = Counter(z['addl_info'][key])






'''
Function loops through the lines of a bundle to gather the needed information
for the summary spreadsheets. 

There are two versions of this function as a form of version control and for testing
performance. The script can take a while to complete.
'''

def processBundle_for_Summaries_original(bundle, contentDict):
    n = 0
    with open(bundle, 'r') as file:
        print(f"File opened and processing started. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        for line in file:
            data = json.loads(line)
            n += 1
            
            info = unpack(data.get('resource', data))
            resType = info.get('resourceType', [])[0]

            # Update categoryCodes
            category_codes = info.get('category_coding_code', [])
            contentDict['categoryCodes'].setdefault(resType, []).extend(category_codes)

            # Update statuses
            statuses = info.get('status', [])
            contentDict['statuses'].setdefault(resType, []).extend(statuses)

            # Update profiles
            profiles = [i.split('/')[-1] for i in info.get('meta_profile', [])]
            contentDict['profiles'].setdefault(resType, []).extend(profiles)

            # Process data elements
            for key in info:
                de = f"{resType}_{key}"
                contentDict['dataElements'].append(de)
                standardCodeSystemUsed = False
                
                if '_system' in key:
                    for sys in info[key]:
                        # Update code systems
                        contentDict['codeSystems'].setdefault(de, []).append(sys)
                        
                        # Check for standardized code systems
                        standardizedCodeSystems = [
                            'http://terminology.hl7.org/CodeSystem/v3-ActCode',
                            'http://terminology.hl7.org/CodeSystem/v3-ActPriority',
                            'http://terminology.hl7.org/CodeSystem/v3-ActUSPrivacyLaw',
                            'http://terminology.hl7.org/CodeSystem/v2-0371',
                            'http://hl7.org/fhir/R4/codesystem-address-type.html',
                            'http://hl7.org/fhir/R4/valueset-address-use.html',
                            'http://hl7.org/fhir/R4/codesystem-administrative-gender.html',
                            'http://terminology.hl7.org/CodeSystem/v3-AdministrativeGender',
                            'http://terminology.hl7.org/CodeSystem/admit-source',
                            'urn:ietf:bcp:47',
                            'http://terminology.hl7.org/CodeSystem/v2-0116',
                            'http://hl7.org/fhir/R4/valueset-bundle-type.html',
                            'urn:oid:2.16.840.1.113883.6.238',
                            'http://terminology.hl7.org/CodeSystem/common-tags',
                            'http://terminology.hl7.org/CodeSystem/condition-category',
                            'http://terminology.hl7.org/CodeSystem/contactentity-type',
                            'http://hl7.org/fhir/contact-point-system',
                            'http://hl7.org/fhir/contact-point-use',
                            'http://terminology.hl7.org/CodeSystem/coverage-class',
                            'http://terminology.hl7.org/CodeSystem/coverage-copay-type',
                            'http://www.ama-assn.org/go/cpt',
                            'http://terminology.hl7.org/CodeSystem/data-absent-reason',
                            'http://hl7.org/fhir/days-of-week',
                            'http://hl7.org/fhir/device-nametype',
                            'http://hl7.org/fhir/device-status',
                            'http://terminology.hl7.org/CodeSystem/diagnosis-role',
                            'http://hl7.org/fhir/diagnostic-report-status',
                            'http://terminology.hl7.org/CodeSystem/diet',
                            'http://terminology.hl7.org/CodeSystem/dose-rate-type',
                            'http://hl7.org/fhir/encounter-location-status',
                            'http://hl7.org/fhir/encounter-status',
                            'http://hl7.org/fhir/event-status',
                            'http://terminology.hl7.org/CodeSystem/list-example-use-codes',
                            'http://hl7.org/fhir/R4/codesystem-device-status.html',
                            'http://terminology.hl7.org/CodeSystem/device-status-reason',
                            'http://hl7.org/fhir/fm-status',
                            'https://www.cms.gov/Medicare/Coding/HCPCSReleaseCodeSets',
                            'urn:oid:2.16.840.1.113883.6.259',
                            'http://hl7.org/fhir/sid/icd-10-cm',
                            'http://www.cms.gov/Medicare/Coding/ICD10',
                            'http://hl7.org/fhir/sid/icd-9-cm',
                            'http://hl7.org/fhir/R4/codesystem-identifier-use.html',
                            'http://hl7.org/fhir/R4/codesystem-link-type.html',
                            'http://hl7.org/fhir/ValueSet/list-mode',
                            'http://hl7.org/fhir/ValueSet/list-status',
                            'http://hl7.org/fhir/R4/codesystem-location-mode.html',
                            'http://hl7.org/fhir/location-status',
                            'http://terminology.hl7.org/CodeSystem/location-physical-type',
                            'http://loinc.org',
                            'http://terminology.hl7.org/CodeSystem/measure-improvement-notation',
                            'http://terminology.hl7.org/CodeSystem/measure-population',
                            'http://hl7.org/fhir/measure-report-status',
                            'http://hl7.org/fhir/measure-report-type',
                            'http://hl7.org/fhir/CodeSystem/medication-status',
                            'http://terminology.hl7.org/CodeSystem/medicationrequest-category',
                            'http://terminology.hl7.org/CodeSystem/medicationrequest-course-of-therapy',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-intent.html',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-status.html',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-status-reason.html',
                            'http://hl7.org/fhir/name-use',
                            'http://terminology.hl7.org/CodeSystem/observation-category',
                            'http://terminology.hl7.org/CodeSystem/referencerange-meaning',
                            'http://hl7.org/fhir/observation-status',
                            'http://terminology.hl7.org/CodeSystem/organization-type',
                            'urn:oid:1.2.36.1.2001.1001.101.104.16592',
                            'http://terminology.hl7.org/CodeSystem/v2-0092',
                            'http://terminology.hl7.org/CodeSystem/v2-0916',
                            'http://hl7.org/fhir/request-intent',
                            'http://hl7.org/fhir/request-priority',
                            'http://hl7.org/fhir/request-status',
                            'http://terminology.hl7.org/CodeSystem/v3-RoleCode',
                            'http://www.nlm.nih.gov/research/umls/rxnorm',
                            'http://terminology.hl7.org/CodeSystem/service-type',
                            'http://snomed.info/sct',
                            'https://nahdo.org/sopt',
                            'http://terminology.hl7.org/CodeSystem/encounter-special-arrangements',
                            'http://terminology.hl7.org/CodeSystem/v2-0493',
                            'http://hl7.org/fhir/specimen-status',
                            'http://terminology.hl7.org/CodeSystem/subscriber-relationship',
                            'http://terminology.hl7.org/CodeSystem/v3-substanceAdminSubstitution',
                            'http://terminology.hl7.org/CodeSystem/v3-TribalEntityUS',
                            'http://unitsofmeasure.org',
                            'http://hl7.org/fhir/udi-entry-type',
                            'http://hl7.org/fhir/us/core/CodeSystem/us-core-category',
                            'https://www.usps.com/',
                            'http://terminology.hl7.org/CodeSystem/v2-0131',
                            'http://terminology.hl7.org/CodeSystem/v2-0203',
                            'http://terminology.hl7.org/CodeSystem/v2-0373',
                            'http://terminology.hl7.org/CodeSystem/v3-EncounterSpecialCourtesy',
                            'http://terminology.hl7.org/CodeSystem/v3-ObservationInterpretation',
                            'http://terminology.hl7.org/CodeSystem/v3-MaritalStatus',
                            'http://terminology.hl7.org/CodeSystem/v3-NullFlavor'
                        ]
                        
                        if sys in standardizedCodeSystems or 'hl7.org' in sys:
                            standardCodeSystemUsed = True
                    
                    if standardCodeSystemUsed:
                        contentDict['Rxnorm_Snomed_Loinc_byResourceType'].setdefault(de, 0)
                        contentDict['Rxnorm_Snomed_Loinc_byResourceType'][de] += 1
            
            # Optional: Limit processing for testing
            # if n > 5000:
            #     break
        
            if n % 250000 == 0:
                print(f"Processed {n/1000000}M lines.")
        
    print(f"Processing Complete for all {n} lines. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    return contentDict


def processBundle_for_Summaries(bundle, contentDict):

    n=0
    with open (bundle, 'r') as file:
            
        print(f"File opened and processing started. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        for line in file:
            data = json.loads(line)
            
            n +=1
            
            info = unpack(data.get('resource', data))
                
            resType = info.get('resourceType', [])[0]
            #resTypes.append(resType)
            
            
            if resType in contentDict['categoryCodes'].keys():
                contentDict['categoryCodes'][resType] += [i for i in info.get('category_coding_code', [])]
            else:
                contentDict['categoryCodes'][resType] = []
                contentDict['categoryCodes'][resType] += [i for i in info.get('category_coding_code', [])]
            
            if resType in contentDict['statuses'].keys():
                contentDict['statuses'][resType] += [i for i in info.get('status', [])]
            else:
                contentDict['statuses'][resType] = []
                contentDict['statuses'][resType] += [i for i in info.get('status', [])]
            
            if resType in contentDict['profiles'].keys():
                contentDict['profiles'][resType] += [i.split('/')[-1] for i in info.get('meta_profile', [])]
            else:
                contentDict['profiles'][resType] = []
                contentDict['profiles'][resType] += [i.split('/')[-1] for i in info.get('meta_profile', [])]
            
            for key in info:
                de = f"{resType}_{key}"
                contentDict['dataElements'].append(de)
                standardCodeSystemUsed = False
                
                if '_system' in key:
                    for sys in info[key]:
                        # Get the code system's frequency tables
                        if de in contentDict['codeSystems'].keys():
                            contentDict['codeSystems'][de].append(sys)
                        else:
                            contentDict['codeSystems'][de] = []
                            contentDict['codeSystems'][de].append(sys)
                        
                        # Determine whether standardized systems are reported
                        standardizedCodeSystems = [
                            'http://terminology.hl7.org/CodeSystem/v3-ActCode',
                            'http://terminology.hl7.org/CodeSystem/v3-ActPriority',
                            'http://terminology.hl7.org/CodeSystem/v3-ActUSPrivacyLaw',
                            'http://terminology.hl7.org/CodeSystem/v2-0371',
                            'http://hl7.org/fhir/R4/codesystem-address-type.html',
                            'http://hl7.org/fhir/R4/valueset-address-use.html',
                            'http://hl7.org/fhir/R4/codesystem-administrative-gender.html',
                            'http://terminology.hl7.org/CodeSystem/v3-AdministrativeGender',
                            'http://terminology.hl7.org/CodeSystem/admit-source',
                            'urn:ietf:bcp:47',
                            'http://terminology.hl7.org/CodeSystem/v2-0116',
                            'http://hl7.org/fhir/R4/valueset-bundle-type.html',
                            'urn:oid:2.16.840.1.113883.6.238',
                            'http://terminology.hl7.org/CodeSystem/common-tags',
                            'http://terminology.hl7.org/CodeSystem/condition-category',
                            'http://terminology.hl7.org/CodeSystem/contactentity-type',
                            'http://hl7.org/fhir/contact-point-system',
                            'http://hl7.org/fhir/contact-point-use',
                            'http://terminology.hl7.org/CodeSystem/coverage-class',
                            'http://terminology.hl7.org/CodeSystem/coverage-copay-type',
                            'http://www.ama-assn.org/go/cpt',
                            'http://terminology.hl7.org/CodeSystem/data-absent-reason',
                            'http://hl7.org/fhir/days-of-week',
                            'http://hl7.org/fhir/device-nametype',
                            'http://hl7.org/fhir/device-status',
                            'http://terminology.hl7.org/CodeSystem/diagnosis-role',
                            'http://hl7.org/fhir/diagnostic-report-status',
                            'http://terminology.hl7.org/CodeSystem/diet',
                            'http://terminology.hl7.org/CodeSystem/dose-rate-type',
                            'http://hl7.org/fhir/encounter-location-status',
                            'http://hl7.org/fhir/encounter-status',
                            'http://hl7.org/fhir/event-status',
                            'http://terminology.hl7.org/CodeSystem/list-example-use-codes',
                            'http://hl7.org/fhir/R4/codesystem-device-status.html',
                            'http://terminology.hl7.org/CodeSystem/device-status-reason',
                            'http://hl7.org/fhir/fm-status',
                            'https://www.cms.gov/Medicare/Coding/HCPCSReleaseCodeSets',
                            'urn:oid:2.16.840.1.113883.6.259',
                            'http://hl7.org/fhir/sid/icd-10-cm',
                            'http://www.cms.gov/Medicare/Coding/ICD10',
                            'http://hl7.org/fhir/sid/icd-9-cm',
                            'http://hl7.org/fhir/R4/codesystem-identifier-use.html',
                            'http://hl7.org/fhir/R4/codesystem-link-type.html',
                            'http://hl7.org/fhir/ValueSet/list-mode',
                            'http://hl7.org/fhir/ValueSet/list-status',
                            'http://hl7.org/fhir/R4/codesystem-location-mode.html',
                            'http://hl7.org/fhir/location-status',
                            'http://terminology.hl7.org/CodeSystem/location-physical-type',
                            'http://loinc.org',
                            'http://terminology.hl7.org/CodeSystem/measure-improvement-notation',
                            'http://terminology.hl7.org/CodeSystem/measure-population',
                            'http://hl7.org/fhir/measure-report-status',
                            'http://hl7.org/fhir/measure-report-type',
                            'http://hl7.org/fhir/CodeSystem/medication-status',
                            'http://terminology.hl7.org/CodeSystem/medicationrequest-category',
                            'http://terminology.hl7.org/CodeSystem/medicationrequest-course-of-therapy',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-intent.html',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-status.html',
                            'http://hl7.org/fhir/R4/codesystem-medicationrequest-status-reason.html',
                            'http://hl7.org/fhir/name-use',
                            'http://terminology.hl7.org/CodeSystem/observation-category',
                            'http://terminology.hl7.org/CodeSystem/referencerange-meaning',
                            'http://hl7.org/fhir/observation-status',
                            'http://terminology.hl7.org/CodeSystem/organization-type',
                            'urn:oid:1.2.36.1.2001.1001.101.104.16592',
                            'http://terminology.hl7.org/CodeSystem/v2-0092',
                            'http://terminology.hl7.org/CodeSystem/v2-0916',
                            'http://hl7.org/fhir/request-intent',
                            'http://hl7.org/fhir/request-priority',
                            'http://hl7.org/fhir/request-status',
                            'http://terminology.hl7.org/CodeSystem/v3-RoleCode',
                            'http://www.nlm.nih.gov/research/umls/rxnorm',
                            'http://terminology.hl7.org/CodeSystem/service-type',
                            'http://snomed.info/sct',
                            'https://nahdo.org/sopt',
                            'http://terminology.hl7.org/CodeSystem/encounter-special-arrangements',
                            'http://terminology.hl7.org/CodeSystem/v2-0493',
                            'http://hl7.org/fhir/specimen-status',
                            'http://terminology.hl7.org/CodeSystem/subscriber-relationship',
                            'http://terminology.hl7.org/CodeSystem/v3-substanceAdminSubstitution',
                            'http://terminology.hl7.org/CodeSystem/v3-TribalEntityUS',
                            'http://unitsofmeasure.org',
                            'http://hl7.org/fhir/udi-entry-type',
                            'http://hl7.org/fhir/us/core/CodeSystem/us-core-category',
                            'https://www.usps.com/',
                            'http://terminology.hl7.org/CodeSystem/v2-0131',
                            'http://terminology.hl7.org/CodeSystem/v2-0203',
                            'http://terminology.hl7.org/CodeSystem/v2-0373',
                            'http://terminology.hl7.org/CodeSystem/v3-EncounterSpecialCourtesy',
                            'http://terminology.hl7.org/CodeSystem/v3-ObservationInterpretation',
                            'http://terminology.hl7.org/CodeSystem/v3-MaritalStatus',
                            'http://terminology.hl7.org/CodeSystem/v3-NullFlavor'
                        ]
                        
                        if sys in standardizedCodeSystems or 'hl7.org' in sys:
                            standardCodeSystemUsed = True
                    
                    if standardCodeSystemUsed:
                        if de in contentDict['Rxnorm_Snomed_Loinc_byResourceType'].keys():
                            contentDict['Rxnorm_Snomed_Loinc_byResourceType'][de] += 1
                        else:
                            contentDict['Rxnorm_Snomed_Loinc_byResourceType'][de] = 0
                            contentDict['Rxnorm_Snomed_Loinc_byResourceType'][de] += 1
                    
            # if n > 5000:
            #      break
        
            if n % 250000 == 0:
                print(f"Processed {n/1000000}M lines.")
        file.close()
    
    print(f"Processing Complete for all {n} lines. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    return contentDict




'''
Function establishes the initial counters that are looped through to create
the dataframes.
'''
def create_Initial_Counters(contentDict):
    initialCounters = {}
    #resCounts = Counter(resTypes)
    initialCounters['elementsCounts'] = Counter(contentDict['dataElements'])
    
    # Helper function to do the counting for common dictionary structures
    def createCounter(dataDict):
        out = {}
        for key in dataDict:
            out[key] = Counter(dataDict[key])
        return out
    
    initialCounters['codeSystemCounts'] = createCounter(contentDict['codeSystems'])
    initialCounters['profilesIncluded'] = createCounter(contentDict['profiles'])
    initialCounters['statusesIncluded'] = createCounter(contentDict['statuses'])
    initialCounters['categoryCodesIncluded'] = createCounter(contentDict['categoryCodes'])
    
    print("Initial Counters Created.")    
    return initialCounters




'''
function creates the initial DFs for storing the counts of information
in the summaries.
'''
def createDFs(contentDict, initialCounters):
    # Create dataframes
    #RTypes = pd.DataFrame(resCounts.items(), columns=['Item', 'Count'])
    DEs = pd.DataFrame(initialCounters['elementsCounts'].items(), columns=['Item', 'Count'])
    standardSysUse_byResEle = pd.DataFrame(contentDict['Rxnorm_Snomed_Loinc_byResourceType'].items(), columns=['Item', 'N_using_RxNorm_SNOMED_LOINC'])
    
    # Need to merge DEs and standardSysUse_byResEle on the 'Item' column.
    df = pd.merge(DEs, standardSysUse_byResEle, on=['Item'], how='outer')
    
    # Function to split the 'Item' column into 'Resource' and 'Element'
    def split_item(item):
        parts = item.split('_', 1)  # Split only at the first underscore
        resource = parts[0]
        element = parts[1] if len(parts) > 1 else None  # Handle cases without an underscore
        return pd.Series([resource, element])
    
    
    df[['Resource', 'Element']] = df['Item'].apply(split_item)
    
    print("Initial dataframe created.")
    
    return df


'''
Function works with the codeSystems in contentDict and initialCounters
to create a dataframe and then merge with the existing dataframe.
'''
def get_codeSystemInfo(initialCounters, df):
    # This provides more detailed information about the systems reported for each data element
    codeSystemInfo = {'Item': [], 'codeSystemUsed': [], 'useCount': []}
    
    for key in initialCounters['codeSystemCounts']:
        for sys in initialCounters['codeSystemCounts'][key]:
            codeSystemInfo['Item'].append(key)
            codeSystemInfo['codeSystemUsed'].append(sys)
            codeSystemInfo['useCount'].append(initialCounters['codeSystemCounts'][key][sys])
    
    codeSystemInfo = pd.DataFrame(codeSystemInfo)
    
    codeSystemInfo = codeSystemInfo.groupby('Item').agg({
        'codeSystemUsed': list,
        'useCount': list}).reset_index()
    
    df = pd.merge(df, codeSystemInfo, on=['Item'], how='outer')
    
    # Any row where 'codeSystemUsed' is not nan but 'N_' is should have 'N_...' set to 0
    df.loc[df['codeSystemUsed'].notna() & df['N_using_RxNorm_SNOMED_LOINC'].isna(), 'N_using_RxNorm_SNOMED_LOINC'] = 0
    
    # Find proportions
    df['proportion_using_RxNorm_SNOMED_LOINC'] = df['N_using_RxNorm_SNOMED_LOINC'] / df['Count']
    df['opposite_proportion'] = 1 - df['proportion_using_RxNorm_SNOMED_LOINC']
    tmp = df[['Resource', 'Element', 'Count', 'N_using_RxNorm_SNOMED_LOINC', 'proportion_using_RxNorm_SNOMED_LOINC', 'opposite_proportion', 'codeSystemUsed', 'useCount']]
    
    
    print("Code System info added to dataframe.")
    return tmp


'''
Function loops through the rest of the InitialCounters items to add their 
information to the summary dataframe.
'''
def getIncluded(dataDict, varname_prefix=None, element=None):
    # name variables
    v_used = f"{varname_prefix}Used"
    v_count = f"{varname_prefix}UseCount"
    
    # initialize output dictionary
    out = {'Resource': [], 'Element': [], v_used: [], v_count: []}
    
    
    for key in dataDict:
        for x in dataDict[key]:
            out['Resource'].append(key)
            out['Element'].append(element)
            out[v_used].append(x)
            out[v_count].append(dataDict[key][x])
    
    out = pd.DataFrame(out)
    out = out.groupby(['Resource', 'Element']).agg({v_used: list, v_count: list}).reset_index()
    
    print(f"{varname_prefix} info added to the dataframe.")
    return out



# =============================================================================
# Run for a single ACH combined file.
# =============================================================================


contentDict = {
        #'resTypes': [],
        'dataElements': [],
        'codeSystems': {},
        'Rxnorm_Snomed_Loinc_byResourceType': {},
        'profiles': {},
        'statuses': {},
        'categoryCodes': {}
    }

mode = 'single' #change this based on which type of file I'm running

if mode == 'single':
    # Individual file processing

    bundle = "//assv-nhsn-blc1/NHSNLink2/reports/working/zzz_Combined Files for Bundle Summaries/Baycare_ACHMonthly_Aug2025_combined.ndjson" 
    '''Add file name before end quotes'''
    #need to tell it the directory here (full filepath) #filedialog.askopenfilename()
    
    facName =  'Baycare'
    #hard code actual facility name here. This is what will appear in the dashboard for facility name #args.facility
    monYear = 'Aug_2025'
    #hard code actual mo/year here # args.month_year  
    
    for_Molly = False #change to False if spreadsheet summary is for Zabrina's team

    if for_Molly: 
        contentDict = processBundle_for_Summaries(bundle, contentDict) # For Molly's dashboard
    else:
        contentDict = processBundle_for_Summaries_original(bundle, contentDict)
    
    initialCounters = create_Initial_Counters(contentDict)
    df = createDFs(contentDict, initialCounters)
    tmp = get_codeSystemInfo(initialCounters, df)
    
    
    tmp1 = getIncluded(initialCounters['profilesIncluded'], 'profiles', 'meta_profile')
    tmp = pd.merge(tmp, tmp1, on=['Resource', 'Element'], how='outer')            
                
    tmp2 = getIncluded(initialCounters['statusesIncluded'], 'status', 'status')
    tmp = pd.merge(tmp, tmp2, on=['Resource', 'Element'], how='outer')   
    
    tmp3 = getIncluded(initialCounters['categoryCodesIncluded'], 'categoryCode', 'category_coding_code')
    df = pd.merge(tmp, tmp3, on=['Resource', 'Element'], how='outer')  
    #tmp = pd.merge(tmp, tmp3, on=['Resource', 'Element'], how='outer')  
    
    
    if for_Molly:
        
        # # Bring in the FHIR DD to the dataframe.
        '''Update this to JW folder'''
        FHIR_DD = 'C:/Users/urb7/OneDrive - CDC/Data Quality Summaries/Bundle_readable_DD_052025 - Copy.xlsx'
        dd = pd.read_excel(FHIR_DD)
        dd.columns = ['Resource', 'Element', 'usedInAnalysesFlag', 'conformanceVerb', 'bindingStrength']
        
        df = pd.merge(tmp, dd, on=['Resource', 'Element'], how='outer' )
        
        print("FHIR data dictionary info added to dataframe.")
        
        # # Any row where 'Count' is nan should get set to 0 - these are the ones that didn't show up in the bundle but are reflected in the data dictionary.
        df.loc[df['Count'].isna(), 'Count'] = 0
    
    
    # =============================================================================
    # # Make sure to adjust this field!!
    # =============================================================================
    if facName == '':
        df['Facility'] = "UC_Davis" # Update manually if needed.
    else:
        df['Facility'] = facName
        
    if monYear == '':
        df['Month_Year'] = "Apr_2025" # Update manually if needed.    
    else:
        df['Month_Year'] = monYear
    print("Bundle meta data added to dataframe.")
    
    
    '''Update this to JW folder'''
    # Write the results to an .xlsx file
    #outdirectory = filedialog.askdirectory()
    outdirectory = "C:/Users/urb7/OneDrive - CDC/Data Quality Summaries"
    outFile = bundle.split('/')[-1].split('.ndjson')[0]
    

    if for_Molly:
        df.to_excel(outdirectory+"/"+outFile+"_data_elements_05222025_forMB.xlsx", index=False)  #<-- For Molly's dashboard if needed
    else:
        # RTypes.to_excel(outdirectory+"/"+outFile+"_resourceTypes.xlsx", index=False)
        df.to_csv(outdirectory+"/"+outFile+"_data_elements_06062025.csv", index=False)    
    
    print(f"Dataframe exported as spreadsheet. Job complete. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")



elif mode == 'batch':
    
    # =============================================================================
    # A batch file summarizes all of the nested folders
    #Use this only for a collection of combined files in the same folder.
    # Likely for just RPS sets of bundles.
    # =============================================================================
    rps_folder = '//assv-nhsn-blc1/NHSNLink2/reports/working/zzz_Combined Files for Bundle Summaries/URMC_RPS_Mar2025'
    
    # Loop through all files in the directory
    for filename in os.listdir(rps_folder):
        file_path = os.path.join(rps_folder, filename)
        
        print(f'     {filename.split("Urmc_rps_")[1].split("_")[0]}')
        
        if filename.split("Urmc_rps_")[1].split("_")[0] == '03062025':
            continue
    
    
        contentDict = {
                #'resTypes': [],
                'dataElements': [],
                'codeSystems': {},
                'Rxnorm_Snomed_Loinc_byResourceType': {},
                'profiles': {},
                'statuses': {},
                'categoryCodes': {}
            }
    
    
        bundle = file_path
    
    
        contentDict = processBundle_for_Summaries(bundle, contentDict)
        initialCounters = create_Initial_Counters(contentDict)
        df = createDFs(contentDict, initialCounters)
        tmp = get_codeSystemInfo(initialCounters, df)
    
    
        tmp1 = getIncluded(initialCounters['profilesIncluded'], 'profiles', 'meta_profile')
        tmp = pd.merge(tmp, tmp1, on=['Resource', 'Element'], how='outer')            
                    
        tmp2 = getIncluded(initialCounters['statusesIncluded'], 'status', 'status')
        tmp = pd.merge(tmp, tmp2, on=['Resource', 'Element'], how='outer')   
        
        tmp3 = getIncluded(initialCounters['categoryCodesIncluded'], 'categoryCode', 'category_coding_code')
        tmp = pd.merge(tmp, tmp3, on=['Resource', 'Element'], how='outer')  
    
    
        # Bring in the FHIR DD to the dataframe.
        FHIR_DD = 'C:/Users/urb7/OneDrive - CDC/Data Quality Summaries/Bundle_readable_DD_052025 - Copy.xlsx'
        dd = pd.read_excel(FHIR_DD)
        dd.columns = ['Resource', 'Element', 'usedInAnalysesFlag', 'conformanceVerb', 'bindingStrength']
        
        df = pd.merge(tmp, dd, on=['Resource', 'Element'], how='outer' )
        
        print("FHIR data dictionary info added to dataframe.")
        
        # Any row where 'Count' is nan should get set to 0 - these are the ones that didn't show up in the bundle but are reflected in the data dictionary.
        df.loc[df['Count'].isna(), 'Count'] = 0
    
    
        # =============================================================================
        # # Make sure to adjust this field!!
        # =============================================================================
        df['Facility'] = 'URMC_rps'
        df['Month_Year'] = f'{filename.split("Urmc_rps_")[1].split("_")[0][0:2]}_{filename.split("Urmc_rps_")[1].split("_")[0][2:4]}_{filename.split("Urmc_rps_")[1].split("_")[0][4:]}'
        print("Bundle meta data added to dataframe.")
    
    
    
        # Write the results to an .xlsx file
        #outdirectory = filedialog.askdirectory()
        outdirectory = "C:/Users/urb7/OneDrive - CDC/Data Quality Summaries/URMC_rps_mar2025"
        outFile = bundle.split('\\')[-1].split('.ndjson')[0]
    
        # RTypes.to_excel(outdirectory+"/"+outFile+"_resourceTypes.xlsx", index=False)
        df.to_excel(outdirectory+"/"+outFile+"_data_elements.xlsx", index=False)
        
        print("Dataframe exported as spreadsheet. Job complete.")


