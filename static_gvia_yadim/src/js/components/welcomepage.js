/**
 * Created by DEV on 9/8/2016.
 */
import React from 'react';
import { PropTypes } from 'react';



class WelcomePage extends React.Component {
    constructor(props) {
        super(props);


    }


    render() {

        return (
           <div className="container">
                <div className="input-field col s12">
                    <h1>Welcome</h1>
                </div>
            </div>
        );
    }
}


export default WelcomePage