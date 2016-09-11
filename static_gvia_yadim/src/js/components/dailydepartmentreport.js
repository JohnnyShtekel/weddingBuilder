import React from 'react';

class DailyDepartmentReport extends React.Component {
    constructor(props) {
        super(props);


    }

    componentDidMount() {

    }


    componentDidUpdate() {

    }


    render() {
        return (
            <div className="container cardContainer animated">
                <div className="card blue-grey darken-1">
                    <div class="card-content white-text" dir="rtl">
                        <span class="card-title">ברוכים הבאים</span>
                        <div className="input-field col s12">
                            <input type="date" class="datepicker"/>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}


export default DailyDepartmentReport