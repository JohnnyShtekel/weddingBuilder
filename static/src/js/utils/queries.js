import React from 'react';

export const  GET_ALL_CONSULERS = "select  FirstName + ' ' + LastName as FullName" +
                                  "from tblWorkers " +
                                  "where active=1";