# ISO/IEC 25012

The Data Quality model represents the grounds where the system for assessing the quality of data products is built on. In a Data Quality model, the main Data Quality characteristics that must be taken into account when assessing the properties of the intended data product are established.

The Quality of a Data Product may be understood as the degree to which data satisfy the requirements defined by the product-owner organization. Specifically, those requirements are the ones that are reflected in the Data Quality model through its characteristics (Accuracy, Completeness, Consistency, Credibility, Currentness, Accessibility...).

The Data Quality model defined in the standard ISO/IEC 25012 is composed of 15 characteristics, shown in the picture bellow:

[![ISO/IEC 25012](https://iso25000.com/images/figures/ISO_25012_en.png "ISO/IEC 25012")](https://iso25000.com/images/figures/ISO_25012_en.png)

The Data Quality characteristics are classified in to main categories:

- **Inherent Data Quality**: Inherent data quality refers to the degree to which quality characteristics of data have the intrinsic potential to satisfy stated and implied needs when data is used under specified conditions. From the inherent point of view, data quality refers to data itself, in particular to:  
    - data domain values and possible restrictions (e.g. business rules governing the quality required for the characteristic in a given application);
    - relationships of data values (e.g. consistency);
    - metadata.
- **System-Dependent Data Quality**: System dependent data quality refers to the degree to which data quality is reached and preserved within a computer system when data is used under specified conditions.
    
    From this point of view data quality depends on the technological domain in which data are used; it is achieved by the capabilities of computer systems' components such as: hardware devices (e.g. to make data available or to obtain the required precision), computer system software (e.g. backup software to achieve recoverability), and other software (e.g. migration tools to achieve portability).
    

## Inherent Data Quality

## Accuracy

The degree to which data has attributes that correctly represent the true value of the intended attribute of a concept or event in a specific context of use.  
It has two main aspects:

- **Syntactic Accuracy**: Syntactic accuracy is defined as the closeness of the data values to a set of values defined in a domain considered syntactically correct.
- **Semantic Accuracy**: Semantic accuracy is defined as the closeness of the data values to a set of values defined in a domain considered semantically correct.

## Completeness

The degree to which subject data associated with an entity has values for all expected attributes and related entity instances in a specific context of use.

## Consistency

The degree to which data has attributes that are free from contradiction and are coherent with other data in a specific context of use. It can be either or both among data regarding one entity and across similar data for comparable entities.

## Credibility

The degree to which data has attributes that are regarded as true and believable by users in a specific context of use. Credibility includes the concept of authenticity (the truthfulness of origins, attributions, commitments).

## Currentness

The degree to which data has attributes that are of the right age in a specific context of use.

## Inherent and System-Dependent Data Quality

## Accessibility

The degree to which data can be accessed in a specific context of use, particularly by people who need supporting technology or special configuration because of some disability.

## Compliance

The degree to which data has attributes that adhere to standards, conventions or regulations in force and similar rules relating to data quality in a specific context of use.

## Confidentiality

The degree to which data has attributes that ensure that it is only accessible and interpretable by authorized users in a specific context of use. Confidentiality is an aspect of information security (together with availability, integrity) as defined in ISO/IEC 13335-1:2004.

## Efficiency

The degree to which data has attributes that can be processed and provide the expected levels of performance by using the appropriate amounts and types of resources in a specific context of use.

## Precision

The degree to which data has attributes that are exact or that provide discrimination in a specific context of use.

## Traceability

The degree to which data has attributes that provide an audit trail of access to the data and of any changes made to the data in a specific context of use.

## Understandability

The degree to which data has attributes that enable it to be read and interpreted by users, and are expressed in appropriate languages, symbols and units in a specific context of use.  
Some information about data understandability are provided by metadata.

## System-Dependent Data Quality

## Availability

The degree to which data has attributes that enable it to be retrieved by authorized users and/or applications in a specific context of use.

## Portability

The degree to which data has attributes that enable it to be installed, replaced or moved from one system to another preserving the existing quality in a specific context of use.

## Recoverability

The degree to which data has attributes that enable it to maintain and preserve a specified level of operations and quality, even in the event of failure, in a specific context of use.