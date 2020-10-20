import React from 'react';
import { shallow } from 'enzyme';
import App from './App';

describe('home testing', ()=> {
  it('', ()=> {
    let app = shallow(<App />)
    expect(app.find('h1').text()).to.equal('React App')
  })
})
