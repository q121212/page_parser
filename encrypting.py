#!/usr/bin/python3
# -*- coding: utf-8 -*-

import pickle 	# for dump & undumps data structures

def encrypt(string, wrt_sets=None):
  '''This function generates a simple line encryption.

  If set any value for wrt_sets then settings file will be updated.
  This case is possible if the string for page_parser will be changed. In this case need to do:
  - 1st step: extract current string with command: decrypt(read_settings())
  - 2nd step: change link in result encrypted string
  - 3rd step: encrypt 2nd step's string by the next code: encrypt(new_string, 'yes')
    This code update settings.ini file and all shoud be work again!... 
  '''

  first = string.split('m')
  second = []
  for i in first:
    second.append(i.split('e'))
  if wrt_sets:
    with open('settings.ini', 'wb', 1000) as file: #  1000 - buffer
      pickle.dump(second[::-1], file)				   # dump data to file
  return second[::-1]


def decrypt(arr):
  '''This function decrypts the encrypted data'''
  first, second = [], []
  for i in arr[::-1]:
      first.append('e'.join(i))
  print('first: ', first)
  second = 'm'.join(first)
  return second


def read_settings():
  '''Function for reading settings from settings.ini file.'''
  with open('settings.ini', 'rb', 1000) as file: #  1000 - buffer
    extracted_data = pickle.load(file)	   # extract data from file
  return extracted_data
