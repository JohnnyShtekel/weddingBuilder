export const  GET_ALL_CONSULERS = "select NumWorker, FirstName + ' ' + LastName as FullName, EMail " +
                                  "from tblWorkers " +
                                  "where EMail IS NOT NULL and active=1";