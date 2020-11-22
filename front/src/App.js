import React, {useEffect} from 'react';
// import axios from 'axios';
import GlobalConfig from './features/reports/globalconfig'
import PageConfig from './features/reports/reportpageconfig'
import ReviewConfig from './features/reports/reviewconfig'

import { useSelector,useDispatch } from "react-redux";
import { loadConfig,updateConfig} from './features/reports/reportsSlice'

// react-router
import {
    BrowserRouter as Router,
        Route,
        Link
} from "react-router-dom";
// material ui
import {
    makeStyles,
    AppBar,
    Button,
    CssBaseline,
    Divider,
    Grid,
    IconButton,
    Toolbar,
    Typography,
} from "@material-ui/core";
import MenuIcon  from '@material-ui/icons/Menu';

const useStyles = makeStyles((theme) => ({
    root: {
        flexGrow: 1,
    },
    appBar: {
        position: 'relative'
    },
    menuButton: {
        marginRight: theme.spacing(2),
    },

    footer: {
        position: 'bottom'
    }
}))


const About = ()=> {
    const classes = useStyles()
    return (
        <Typography variant={'h6'}>

            <h3>Volkswagen CVI Report Automation App.</h3>
            <Divider />
            <p className={classes.right}>Developed by Volkswagen ITP dev team. Oct. 2020.</p>
        </Typography>
    )
}
function App() {
    const classes = useStyles();
    const config = useSelector(state=> state.reports.config )
    const dispatch = useDispatch()

    const handleSave = async ()=> {
        // const url = "http://localhost:3002/config"
        const url = "http://localhost:5000/config"
        await fetch(url, {
            method: 'POST',
            body: JSON.stringify(config),
            headers: new Headers({
                'Content-Type': 'application/json'
            })
        })
        console.log(config)
    }
    useEffect(async () => {
        const fetchConfig = async ()=> {
            await dispatch(loadConfig())
        }
        await fetchConfig()
    },[])

    return (
      <Router>
          <CssBaseline />
          <header>
              <div className={classes.root}>
                  <AppBar position='static' variant="dense">
                      <Toolbar>
                          <IconButton edge="start" color="inherit" aria-label="menu">
                              <MenuIcon />
                          </IconButton>
                          <Typography variant='subtitle1'>VGC CVI Report Configuration Page</Typography>
                          <div className={classes.root}></div>
                          <Button
                              className={classes.menuButton}
                              color="inherit"
                              onClick={handleSave}
                          >Save & Run</Button>
                      </Toolbar>
                  </AppBar>
              </div>
          </header>
          <main>
            <nav>
                  <Grid container>
                      <Grid item xs={3} sm={2}>
                          <Link to="/">Global Config</Link>
                      </Grid>
                      <Grid item xs={3} sm={2}>
                          <Link to="/page">Slides Config</Link>
                      </Grid>
                      <Grid item xs={3} sm={2}>
                          <Link to="/review">Review config</Link>
                      </Grid>
                      <Grid item xs={3} sm={2}>
                          <Link to="/about">About</Link>
                      </Grid>
                  </Grid>
              </nav>
              <hr />
              <Route>
                  <Route exact path="/"><GlobalConfig config={config}/></Route>
                  <Route path="/page"><PageConfig config={config}/></Route>
                  <Route path="/review"><ReviewConfig config={config}/></Route>
                  <Route path="/about"><About /></Route>
              </Route>
          </main>
          <footer className={classes.footer}>
              <h3>CVI app</h3>
          </footer>
      </Router>
  );
}
export default App;
