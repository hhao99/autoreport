import React, {useState} from 'react';
import {useDispatch, useSelector} from "react-redux";
import {
    makeStyles,
    Divider,
    Grid,
    Paper,
    Tab,
    Tabs

} from '@material-ui/core'

import MultipleSelect from './multipleselect'
import SingleSelect from './select'
import {updateConfig} from "./reportsSlice";

const TabPanel = ({value,index,children, ...others}) => {

    return(
        <div role='tabpanel'
            hidden={value !== index}
             id={`simple-panel-${index}`}
             {...others}
        >
            {children}
        </div>
    )

}
const SlideConfig = ({ config, slides_index}) => {
    const dispatch = useDispatch()
    const slide = config.slides[slides_index]
    console.log("slides index: " + slides_index)
    console.log(config)

    const handleChange = (e,field,slide_index) => {
        let s = slide
        if( field === 'filter') {
            let filters = s.filters.map( (filter,index) => {
                if( index === slide_index) {
                    return {...filter, selected: e.target.value}
                }
                return filter
            })
            s = {...slide,filters: filters}
        }
        else if( field === 'group_by') {
            s = {...slide, group_by: {...slide.group_by, selected: e.target.value}}
        }
        else {
            s = {...slide,
                compute_methods: { ...slide.compute_methods,
                    divided_by: {...slide.compute_methods.divided_by, selected: e.target.value}}}
        }
        let slides = config.slides.map( (slide,index) => {
            if( index === slides_index ) return s
            else return slide
        })
        let conf = {...config,slides}
        dispatch(updateConfig(conf))
    }
    return (
        <React.Fragment>
            <Paper>
                <h3>Local Filter for slide {slide.name}</h3>
                <Grid container spacing={3}>
                    {slide.filters.map ( (filter, index) =>
                        <Grid item xs={3}>
                            <MultipleSelect filter={filter} onChange={e=>handleChange(e,'filter',index)} />
                        </Grid>
                    )}
                </Grid>
                <Divider />
                <h3>Group By config</h3>
                <MultipleSelect filter={slide.group_by} onChange={e=> handleChange(e,'group_by', 0)} />
                <Divider />
                <h3>Compution methods config</h3>
                <MultipleSelect filter={slide.compute_methods.divided_by} onChange={e=> handleChange(e,'divided_by',0)} />

            </Paper>
        </React.Fragment>
    )
}

const  useStyles = makeStyles( (theme) => ({
    tabs: {
        borderRight: `1px solid ${theme.palette.divider}`,
    },
}))
export default ({index})=> {
    const classes = useStyles()
    const state = useSelector( state => state.reports)
    const config = state.config
    const [tab,setTab] = useState(0)

    return (
        <>
        <Paper>
            <Grid container spacing={3}>
                <Grid item xs={2}>
                    <Tabs value={tab}
                          className={classes.tabs}
                          scrollButtons="auto"
                          orientation="vertical"
                          variant="scrollable"
                          onChange={(e,v)=> setTab(v)}>
                        {config.slides.map( (slide,index) => <Tab disabled={!slide.included} label={slide.name} />) }
                    </Tabs>
                </Grid>
                <Grid item xs={10}>
                    {config.slides.map( (slide,index) => {
                        return (
                            <TabPanel index={index} value={tab}>
                                <SlideConfig config={config} slides_index={index} />
                            </TabPanel>
                        )})}
                </Grid>
            </Grid>
        </Paper>
        </>
    )
}