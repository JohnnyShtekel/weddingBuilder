import React from 'react';
import Entry from '../components/entry';
import Nav from '../components/nav';

class App extends React.Component {
    constructor(props) {
        super(props);

    }


    render() {
        return (
            <div>
                <Nav/>
                <Entry/>
            </div>
        )

    }

}

export default App


