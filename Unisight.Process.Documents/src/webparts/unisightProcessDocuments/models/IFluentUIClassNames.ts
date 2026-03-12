import { mergeStyleSets } from '@fluentui/react/lib/Styling';

export const classNames = mergeStyleSets({
    fileIconHeaderIcon: {
      padding: 0,
      fontSize: '16px',
    },
    fileIconCell: {
      textAlign: 'center',
      selectors: {
        '&:before': {
          content: '.',
          display: 'inline-block',
          verticalAlign: 'middle',
          height: '100%',
          width: '0px',
          visibility: 'hidden',
        },
      },
    },
    fileIconImg: {
      verticalAlign: 'middle',
      maxHeight: '16px',
      maxWidth: '16px',
    },
    controlWrapper: {
      display: 'flex',
      flexWrap: 'wrap',
      margin: "0 0 -2em 0"
    },
    controlWrapper2: {
        display: 'flex',
        flexWrap: 'wrap',
        textAlign:'right',
        margin: '-40pt 0 0 275pt'
      },
    exampleToggle: {
      display: 'inline-block',
      marginBottom: '10px',
      marginRight: '30px',
    },
    selectionDetails: {
      marginBottom: '20px',
    },
  });
  export const controlStyles = {
    root: {
      margin: '0 30px 20px 0',
      maxWidth: '300px',
    },
  };