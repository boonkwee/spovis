# -*- coding: utf-8 -*-
"""
Created on Wed Jun 26 09:40:24 2024

@author: chanboonkwee
"""
import hashlib
import string
import random
import numpy as np
# from datetime import datetime

def hash_sha256(raw_str:str='', verbose:bool=False, salt:bool=False, salt_length:int=12) -> str:
    ts = ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(salt_length))
    if raw_str in [None, np.nan]:
        raw_str = b''
    elif isinstance(raw_str, str):
        raw_str = raw_str.encode('utf-8')
    result = hashlib.sha256(ts.encode('utf-8') + raw_str if salt else raw_str)
    output_str = result.hexdigest()
    if verbose:
        print(f"{output_str} - {raw_str[:5]}")
    return output_str


def hash_sha256_randsalt(raw_str:str='', salt_length:int=12, verbose:bool=False) -> str:
    ts = ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(salt_length))
    # ts = datetime.today().strftime('%Y-%m-%dT%H:%M:%S.%fZ')
    if raw_str in [None, np.nan]:
        raw_str = b''
    elif isinstance(raw_str, str):
        raw_str = raw_str.encode('utf-8')
    result = hashlib.sha256(ts.encode('utf-8') + raw_str)
    output_str = result.hexdigest()
    if verbose:
        print(f"{output_str} - {raw_str[:17]}")
    return output_str


if __name__ == '__main__':
    blank_hash = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    blank_hsh = hash_sha256('')
    non_blank_hsh = hash_sha256('', salt=True)
    print(f"Blank hash : {blank_hsh}")
    print(f"Salted hash: {non_blank_hsh}")
    assert(blank_hsh == blank_hash)
    assert(non_blank_hsh != blank_hash)
    # print('ok')