import * as React from 'react';
import styles from './Birthday.module.scss';
import { IBirthdayProps } from './IBirthdayProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IBaseButtonState } from 'office-ui-fabric-react/lib/Button';

/* export interface IBirthdayState{    
  items:[    
        {    
          "Title": "",
          "EmailId": "",
          "Birth Date": Date,
          "Joining Date": Date
        }]    
} */

export default class Birthday extends React.Component<IBirthdayProps, {}> {
  

  /* public constructor(props: IBirthdayProps, state: IBirthdayState){    
    super(props);    
    this.state = {    
      items: [    
        {    
          "Title": "",
          "EmailId": "",
          "Birth Date": new Date(), 
          "Joining Date": new Date()
        }    
      ]    
    };    
  } */

  /* public componentDidMount(){    
    var reactHandler = this;    
    jquery.ajax({    
        url: `${this.props.siteurl}/_api/web/lists/getbytitle('EmployeeMaster')/items`,    
        type: "GET",    
        headers:{'Accept': 'application/json; odata=verbose;'},    
        success: function(resultData) {    
          reactHandler.setState({    
            items: resultData.d.results    
          }); 
          debugger;   
        },    
        error : function(jqXHR, textStatus, errorThrown) {    
        }    
    });    
  }   */  

  /* public render(): React.ReactElement<IBirthdayProps> { 
    
    /* return (    
   
       <div className={styles.birthday} > 
       <div className={styles.container}>      
            <div className={styles.description}> <h1>Birthday/Anniversary</h1> </div> 
             {this.state.items.map(function(item,key){    
                 
               return (<div className={styles.row} key={key}> 
                     
                   <div className={styles.column}>{item.Title}</div>  
                   <div className={styles.column}><a href="">{item.EmailId}</a></div>
                   
                 </div>);    
             })}                        
           </div>  
       </div>    
   ); */
  
}


