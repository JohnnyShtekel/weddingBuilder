import React from 'react';
import * as quries from '../utils/queries'
import Select from 'react-select';

class AutoComplete extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            workers: []
        };
        this.workers = []

    }


    componentDidMount() {
        $(document).ready(function () {
            $('select').material_select();
        });
        //$.ajax({
        //    url: 'http://52.41.181.217/esg-sql-service/search/',
        //    type: 'POST',
        //    data: {query: quries.GET_ALL_CONSULERS},
        //    dataType: 'json',
        //    success: function (data) {
        //        this.setState({workers: data}, function () {
        //        });
        //
        //    }.bind(this)
        //});

          $('#select-counslers').selectize({
              persist: false,
              maxItems: null,
              valueField: 'email',
              labelField: 'name',
              searchField: ['name', 'email'],
              options: [
                  {email: 'brian@thirdroute.com', name: 'Brian Reavis'},
                  {email: 'nikola@tesla.com', name: 'Nikola Tesla'},
                  {email: 'someone@gmail.com'}
              ],
          });

    }


        componentWillMount()
        {


            //$.ajax({
            //    url: 'http://52.41.181.217/esg-sql-service/search/',
            //    type: 'POST',
            //    data: {query: quries.GET_ALL_CONSULERS},
            //    dataType: 'json',
            //    success: function (data) {
            //
            //        this.workers = data;
            //
            //    }.bind(this)
            //});


        }


        render()
        {
            return (

                <div className="col s4 offset-s4">
                    <i className="large material-icons col s4 offset-s4">perm_identity</i>
                    <h1>נא לבחור עובד</h1>
                    <select dir="ltr" id="select-counslers">
                    </select>
                </div>

            );
        }
    }


    export
    default
    AutoComplete